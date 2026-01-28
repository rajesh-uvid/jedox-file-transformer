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

def process_zip_recursive(input_bytes, search_terms, replacements=None, prefix="", flatten_to_zip=None):
    in_buffer = BytesIO(input_bytes)
    out_buffer = BytesIO()
    found_items = []

    with zipfile.ZipFile(in_buffer, 'r') as zin:
        file_list = zin.namelist()
        
        # LOGIC: Find Report Name from sibling item.xml
        # We look for item.xml at the current level to name the current archive
        current_level_report_name = prefix.split(" > ")[-1] if prefix else "Root Archive"
        if "item.xml" in file_list:
            try:
                with zin.open("item.xml") as f:
                    tree = ET.parse(f)
                    # Look for <info><name>Reports</name></info>
                    name_tag = tree.find(".//info/name")
                    if name_tag is not None and name_tag.text:
                        current_level_report_name = name_tag.text
            except:
                pass

        zout = zipfile.ZipFile(out_buffer, 'w', zipfile.ZIP_DEFLATED) if flatten_to_zip is None else None
        
        for item in zin.infolist():
            filename = item.filename
            full_location = f"{prefix} > {filename}" if prefix else filename
            
            with zin.open(item) as f:
                content = f.read()

            # Recursive Drill into .wss
            if filename.lower().endswith('.wss'):
                inner_data, inner_logs = process_zip_recursive(content, search_terms, replacements, prefix=full_location, flatten_to_zip=flatten_to_zip)
                content = inner_data
                found_items.extend(inner_logs)
            
            # Content Search
            elif filename.lower().endswith(('.xml', '.rels', '.pb', '.json', '.txt')):
                try:
                    text = content.decode('utf-8', errors='ignore')
                    
                    for term in search_terms:
                        hits = text.lower().count(term.lower())
                        
                        if hits > 0 or (replacements and f"{full_location}||{term}" in replacements):
                            # Sheets are internal to the .wss, Report Name is from the parent's item.xml
                            sheets = get_sheet_names(zin) if "xl/workbook.xml" in file_list else "N/A"
                            
                            found_items.append({
                                "Report Name": current_level_report_name,
                                "Sheet Names": sheets,
                                "Elaborated Location": full_location,
                                "Search Word": term,
                                "Hits": hits,
                                "Full Content": text,
                                "Replacement Word": ""
                            })
                            
                            if replacements:
                                rep_key = f"{full_location}||{term}"
                                if rep_key in replacements:
                                    new_val = replacements[rep_key]
                                    if new_val.startswith("MANUAL_EDIT_MARKER:"):
                                        text = new_val.replace("MANUAL_EDIT_MARKER:", "")
                                    elif new_val.strip():
                                        text = re.compile(re.escape(term), re.IGNORECASE).sub(new_val, text)
                                    content = text.encode('utf-8')
                except:
                    pass

                if flatten_to_zip is not None:
                    flat_name = full_location.replace(" > ", "/").replace(" ", "_")
                    flatten_to_zip.writestr(flat_name, content)

            if zout:
                zout.writestr(item, content)
        if zout: zout.close()
            
    return out_buffer.getvalue(), found_items

# --- STREAMLIT UI ---
st.title("🚀 Jedox Enterprise Transformer")

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
        edited_df = st.data_editor(
            df[[
                # "Report Name",
                 "Sheet Names", "Elaborated Location", "Search Word", "Hits", "Replacement Word", "MapKey"]], 
            use_container_width=True, hide_index=True,
            column_config={
                "MapKey": None,
                # "Report Name": st.column_config.TextColumn(disabled=True),
                "Sheet Names": st.column_config.TextColumn(disabled=True),
                "Elaborated Location": st.column_config.TextColumn(disabled=True),
                "Search Word": st.column_config.TextColumn(disabled=True),
                "Hits": st.column_config.NumberColumn(disabled=True),
            }
        )

        st.divider()
        st.subheader("📄 Preview & Manual Edit")
        selected_loc = st.selectbox("Select file:", df["Elaborated Location"].unique())
        raw_text = df[df["Elaborated Location"] == selected_loc]["Full Content"].values[0]

        # In-file search helper
        local_find = st.text_input("Find inside this XML:")
        if local_find:
            st.caption(f"Occurrences: {raw_text.lower().count(local_find.lower())}")

        new_text = st.text_area("XML Editor:", value=format_xml(raw_text), height=400)
        if st.button("Save Manual Edit"):
            st.session_state.manual_edits[selected_loc] = "MANUAL_EDIT_MARKER:" + new_text
            st.toast("Saved!")

        st.divider()
        if st.button("🚀 Submit & Build"):
            final_reps = dict(zip(edited_df["MapKey"], edited_df["Replacement Word"]))
            for loc, val in st.session_state.manual_edits.items():
                for term in search_terms:
                    final_reps[f"{loc}||{term}"] = val
            
            with st.spinner("Rebuilding structure..."):
                std_bytes, _ = process_zip_recursive(uploaded_file.getvalue(), search_terms, final_reps)
                flat_buf = BytesIO()
                with zipfile.ZipFile(flat_buf, 'w', zipfile.ZIP_DEFLATED) as fz:
                    process_zip_recursive(uploaded_file.getvalue(), search_terms, final_reps, flatten_to_zip=fz)
                
                st.success("Rebuild Complete!")
                c1, c2 = st.columns(2)
                with c1: st.download_button("📥 Original Format", std_bytes, f"mod_{uploaded_file.name}")
                with c2: st.download_button("📥 Pure XML Zip", flat_buf.getvalue(), "extract.zip")