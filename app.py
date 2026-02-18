import streamlit as st
import zipfile
import pandas as pd
import re
import xml.dom.minidom
from io import BytesIO
import xml.etree.ElementTree as ET

# --- APP CONFIGURATION ---
st.set_page_config(page_title="Jedox Pro Transformer", layout="wide")

def format_xml(xml_string):
    try:
        dom = xml.dom.minidom.parseString(xml_string)
        return dom.toprettyxml(indent="  ")
    except Exception:
        return xml_string

def get_sheet_names(zin):
    """Parses workbook.xml to get tab labels."""
    try:
        if "xl/workbook.xml" in zin.namelist():
            with zin.open("xl/workbook.xml") as f:
                tree = ET.parse(f)
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                sheets = tree.findall(".//main:sheet", ns)
                return ", ".join([s.get("name") for s in sheets])
    except:
        pass
    return ""

def process_zip_recursive(input_bytes, search_terms, replacement_map=None, prefix="", flatten_to_zip=None, inherited_readable_path=""):
    in_buffer = BytesIO(input_bytes)
    out_buffer = BytesIO()
    found_items = []

    with zipfile.ZipFile(in_buffer, 'r') as zin:
        file_list = zin.namelist()
        
        # LOGIC: Pre-scan for all item.xml files to map directories to Report Names
        dir_to_report_name = {}
        for fname in file_list:
            if fname.endswith("item.xml"):
                try:
                    with zin.open(fname) as f:
                        tree = ET.parse(f)
                        name_tag = tree.find(".//info/name")
                        if name_tag is not None and name_tag.text:
                            # Store with directory path as key
                            # "item.xml" -> ""
                            # "folder/item.xml" -> "folder"
                            dname = fname.rpartition("/")[0]
                            dir_to_report_name[dname] = name_tag.text
                except Exception:
                    pass

        # Determine Root Name for this archive context
        # Priority: 1. Root item.xml, 2. Passed prefix (parent), 3. Default
        current_level_report_name = dir_to_report_name.get("", prefix.split(" > ")[-1] if prefix else "Root Archive")

        # Use None for zout if we are strictly flattening to another zip
        zout = None
        if flatten_to_zip is None:
            zout = zipfile.ZipFile(out_buffer, 'w', zipfile.ZIP_DEFLATED)
        
        try:
            for item in zin.infolist():
                filename = item.filename
                full_location = f"{prefix} > {filename}" if prefix else filename
                
                with zin.open(item) as f:
                    content = f.read()

                # Recursive Drill into .wss
                if filename.lower().endswith('.wss'):
                    # Pass the report name of the CURRENT .wss file down to its children
                    # We need to compute the readable path for THIS .wss file to pass it down
                    path_parts_wss = filename.split('/')[:-1] # directory of the .wss
                    friendly_parts_wss = []
                    current_check_wss = ""
                    for part_wss in path_parts_wss:
                        current_check_wss = f"{current_check_wss}/{part_wss}" if current_check_wss else part_wss
                        if current_check_wss in dir_to_report_name:
                            friendly_parts_wss.append(dir_to_report_name[current_check_wss])
                    
                    local_readable = " > ".join(friendly_parts_wss)
                    next_inherited_path = inherited_readable_path
                    if local_readable:
                        next_inherited_path = f"{inherited_readable_path} > {local_readable}" if inherited_readable_path else local_readable

                    inner_data, inner_logs = process_zip_recursive(
                        content, 
                        search_terms, 
                        replacement_map=replacement_map, 
                        prefix=full_location, 
                        flatten_to_zip=flatten_to_zip,
                        inherited_readable_path=next_inherited_path
                    )
                    content = inner_data
                    found_items.extend(inner_logs)
                
                # Content Search & Replace
                elif filename.lower().endswith(('.xml', '.rels', '.pb', '.json', '.txt')):
                    try:
                        text = content.decode('utf-8', errors='ignore')
                        
                        # Determine active terms for this file: Global Search Terms + Specific Replacements
                        file_specific_reps = replacement_map.get(full_location, {}) if replacement_map else {}
                        active_terms = sorted(list(set(search_terms) | set(file_specific_reps.keys())))

                        file_modified = False
                        
                        for term in active_terms:
                            # Count hits (case-insensitive)
                            hits = text.lower().count(term.lower())
                            
                            replacement_val = file_specific_reps.get(term)
                            
                            # Check if we should log/act
                            if hits > 0 or replacement_val:
                                sheets = get_sheet_names(zin) if "xl/workbook.xml" in file_list else "N/A"
                                
                                # Build logical path for this file based on folder mappings
                                # Check every level of the path: segments/files/group/node ...
                                # If a level has a mapped name, include it in the path.
                                path_parts = filename.split('/')[:-1] # Exclude filename itself
                                friendly_parts = []
                                current_check = ""
                                for part in path_parts:
                                    current_check = f"{current_check}/{part}" if current_check else part
                                    if current_check in dir_to_report_name:
                                        friendly_parts.append(dir_to_report_name[current_check])
                                
                                local_friendly_path = " > ".join(friendly_parts)
                                
                                # Combine with inherited path (from parent zip)
                                full_report_path = inherited_readable_path
                                if local_friendly_path:
                                    full_report_path = f"{inherited_readable_path} > {local_friendly_path}" if inherited_readable_path else local_friendly_path
                                
                                # Fallback if path is empty (e.g. root files without mapping)
                                if not full_report_path:
                                    full_report_path = "Root"

                                # Sanitize Elaborated Location
                                sanitized_full_location = full_location
                                if sanitized_full_location.startswith("Root Archive > "):
                                    sanitized_full_location = sanitized_full_location[len("Root Archive > "):]

                                found_items.append({
                                    "Report Name": full_report_path,
                                    "Sheet Names": sheets,
                                    "Elaborated Location": sanitized_full_location,
                                    "Search Word": term,
                                    "Hits": hits,
                                    "Full Content": text, # Note: This stores text state at this moment
                                    "Replacement Word": replacement_val if replacement_val else ""
                                })
                                    
                                # Apply Replacement
                                if replacement_val:
                                    if replacement_val.startswith("MANUAL_EDIT_MARKER:"):
                                        text = replacement_val.replace("MANUAL_EDIT_MARKER:", "")
                                        file_modified = True
                                    elif replacement_val.strip():
                                        # Regex replace for case-insensitive match
                                        pattern = re.compile(re.escape(term), re.IGNORECASE)
                                        text = pattern.sub(replacement_val, text)
                                        file_modified = True
                        
                        if file_modified:
                            content = text.encode('utf-8')
                    except Exception:
                        pass # Fail safe for decoding issues

                    if flatten_to_zip is not None:
                        flat_name = full_location.replace(" > ", "/").replace(" ", "_")
                        flatten_to_zip.writestr(flat_name, content)

                if zout:
                    zout.writestr(item, content)
        finally:
            if zout:
                zout.close()
            
    return out_buffer.getvalue(), found_items

# --- STREAMLIT UI ---
st.title("🚀 Jedox Report Parser")

uploaded_file = st.file_uploader("Upload Jedox .pb or .wss", type=["pb", "wss"])

if uploaded_file:
    if 'matches' not in st.session_state or st.session_state.get('file_id') != uploaded_file.name:
        st.session_state.matches = []
        st.session_state.file_id = uploaded_file.name
        st.session_state.manual_edits = {}

    search_input = st.text_input("Search terms (comma-separated):", value="PALO.DATAC, PALO.DATA")
    search_terms = [t.strip() for t in search_input.split(",") if t.strip()]

    if st.button("🔍 Deep Scan (Include Metadata)"):
        with st.spinner("Extracting Report Names from item.xml and Sheet Names..."):
            _, logs = process_zip_recursive(uploaded_file.getvalue(), search_terms)
            st.session_state.matches = logs

    if st.session_state.matches:
        df = pd.DataFrame(st.session_state.matches)
        df["MapKey"] = df["Elaborated Location"] + "||" + df["Search Word"]

        st.subheader("📍 Found Locations")
        
        # Download Analysis Report
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Analysis')
            return output.getvalue()

        col_act1, col_act2 = st.columns([1, 4])
        with col_act1:
            st.download_button(
                label="📥 Download Analysis Excel",
                data=to_excel(df),
                file_name=f"jedox_analysis_{uploaded_file.name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        edited_df = st.data_editor(
            df[[ "Report Name", "Elaborated Location", "Search Word", "Hits", "Replacement Word", "MapKey"]], 
            use_container_width=True, hide_index=True,
            column_config={
                "MapKey": None,
                "Report Name": st.column_config.TextColumn(disabled=True),
                "Sheet Names": None,
                "Elaborated Location": st.column_config.TextColumn(disabled=True),
                "Search Word": st.column_config.TextColumn(disabled=True),
                "Hits": st.column_config.NumberColumn(disabled=True),
            }
        )

        st.divider()
        st.subheader("📄 Preview & Manual Edit")
        selected_loc = st.selectbox("Select file:", df["Elaborated Location"].unique())
        raw_text = df[df["Elaborated Location"] == selected_loc]["Full Content"].values[0]



        new_text = st.text_area("XML Editor:", value=format_xml(raw_text), height=400)
        if st.button("Save Manual Edit"):
            st.session_state.manual_edits[selected_loc] = "MANUAL_EDIT_MARKER:" + new_text
            st.toast("Saved!")

        st.divider()
        if st.button("🚀 Submit & Build"):
            # Prepare Replacement Map: { Location: { Term: Replacement } }
            replacement_map = {}
            
            # 1. Gather replacements from the Data Editor table
            # Retrieve the latest state from the data editor
            final_reps_flat = dict(zip(edited_df["MapKey"], edited_df["Replacement Word"]))
            
            for key, val in final_reps_flat.items():
                if val and val.strip():
                    loc, term = key.split("||")
                    if loc not in replacement_map:
                        replacement_map[loc] = {}
                    replacement_map[loc][term] = val
            
            # 2. Gather replacements from Manual Edits
            # Ensure we attach manual edits to at least one term so they are picked up in the loop.
            # We iterate over current search_terms to ensure coverage.
            for loc, val in st.session_state.manual_edits.items():
                # OVERRIDE: If manual edit exists, discard any individual column replacements for this file
                replacement_map[loc] = {}

                # If there are search terms, associate with all of them to be safe
                if search_terms:
                    for term in search_terms:
                        replacement_map[loc][term] = val
                else: 
                     # If no search terms, use a dummy term to ensure loop entry
                    replacement_map[loc]["__MANUAL__"] = val

            with st.spinner("Rebuilding structure..."):
                std_bytes, _ = process_zip_recursive(uploaded_file.getvalue(), search_terms, replacement_map=replacement_map)
                
                flat_buf = BytesIO()
                with zipfile.ZipFile(flat_buf, 'w', zipfile.ZIP_DEFLATED) as fz:
                    process_zip_recursive(uploaded_file.getvalue(), search_terms, replacement_map=replacement_map, flatten_to_zip=fz)
                
                st.success("Rebuild Complete!")
                c1, c2 = st.columns(2)
                with c1: st.download_button("📥 Original Format", std_bytes, f"mod_{uploaded_file.name}")
                with c2: st.download_button("📥 Pure XML Zip", flat_buf.getvalue(), "extract.zip")

    if st.sidebar.button("⚠️ Clear Session"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()