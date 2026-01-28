import streamlit as st
import zipfile
import pandas as pd
import re
import xml.dom.minidom
from io import BytesIO

# --- APP CONFIGURATION ---
st.set_page_config(page_title="Jedox Pro Transformer", layout="wide")

def format_xml(xml_string):
    try:
        dom = xml.dom.minidom.parseString(xml_string)
        return dom.toprettyxml(indent="  ")
    except Exception:
        return xml_string

# --- RECURSIVE ENGINE ---
def process_zip_recursive(input_bytes, search_term, replacements=None, prefix="", flatten_to_zip=None):
    in_buffer = BytesIO(input_bytes)
    out_buffer = BytesIO()
    found_items = []

    with zipfile.ZipFile(in_buffer, 'r') as zin:
        zout = zipfile.ZipFile(out_buffer, 'w', zipfile.ZIP_DEFLATED) if flatten_to_zip is None else None
        
        for item in zin.infolist():
            filename = item.filename
            full_location = f"{prefix} > {filename}" if prefix else filename
            
            with zin.open(item) as f:
                content = f.read()

            if filename.lower().endswith('.wss'):
                inner_data, inner_logs = process_zip_recursive(content, search_term, replacements, prefix=full_location, flatten_to_zip=flatten_to_zip)
                content = inner_data
                found_items.extend(inner_logs)
            
            elif filename.lower().endswith(('.xml', '.rels', '.pb', '.json', '.txt')):
                try:
                    text = content.decode('utf-8', errors='ignore')
                    hits = text.lower().count(search_term.lower())
                    
                    if hits > 0 or (replacements and full_location in replacements):
                        found_items.append({
                            "Elaborated Location": full_location,
                            "Search Word": search_term,
                            "Hits": hits,
                            "Full Content": text,
                            "Replacement Word": ""
                        })
                        
                        # Apply table-based replacement
                        if replacements and full_location in replacements:
                            new_val = replacements[full_location]
                            # If it's a direct manual edit of the whole file
                            if new_val.startswith("MANUAL_EDIT_MARKER:"):
                                text = new_val.replace("MANUAL_EDIT_MARKER:", "")
                            # If it's a search/replace word
                            elif new_val.strip():
                                text = re.compile(re.escape(search_term), re.IGNORECASE).sub(new_val, text)
                            
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

# --- MAIN UI ---
st.title("🛠️ Jedox Pro: Advanced Transformer")

uploaded_file = st.file_uploader("Upload Jedox Archive (.pb or .wss)", type=["pb", "wss"])

if uploaded_file:
    if 'matches' not in st.session_state or st.session_state.get('file_id') != uploaded_file.name:
        st.session_state.matches = []
        st.session_state.file_id = uploaded_file.name
        st.session_state.manual_edits = {}

    search_word = st.text_input("Enter word to search across all files:", value="PALO.DATAC")

    if st.button("🔍 Deep Scan Structure"):
        with st.spinner("Analyzing recursive structure..."):
            _, logs = process_zip_recursive(uploaded_file.getvalue(), search_word)
            st.session_state.matches = logs

    if st.session_state.matches:
        st.divider()
        df = pd.DataFrame(st.session_state.matches)

        # --- SECTION 1: FOUND LOCATIONS ---
        st.subheader("📍 Found Locations")
        # Column configuration to make only Replacement Word editable
        edited_df = st.data_editor(
            df[["Elaborated Location", "Search Word", "Hits", "Replacement Word"]], 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Elaborated Location": st.column_config.TextColumn(disabled=True),
                "Search Word": st.column_config.TextColumn(disabled=True),
                "Hits": st.column_config.NumberColumn(disabled=True),
                "Replacement Word": st.column_config.TextColumn(disabled=False)
            }
        )

        # --- SECTION 2: PREVIEW & MANUAL EDIT ---
        st.divider()
        st.subheader("📄 File Preview & Manual Edit")
        
        selected_loc = st.selectbox("Select file to preview or edit:", df["Elaborated Location"].unique())
        raw_text = df[df["Elaborated Location"] == selected_loc]["Full Content"].values[0]

        # In-file search feature
        local_search = st.text_input(f"Find word inside {selected_loc.split('>')[-1]}:")
        if local_search:
            local_hits = raw_text.lower().count(local_search.lower())
            st.caption(f"Found {local_hits} occurrences of '{local_search}' in this file.")

        # Manual Editor Area
        st.info("💡 You can manually edit the XML code below. Changes will be saved when you build the file.")
        new_manual_text = st.text_area("Edit File Content:", value=format_xml(raw_text), height=400)
        
        if st.button("Save Manual Edit for this File"):
            st.session_state.manual_edits[selected_loc] = "MANUAL_EDIT_MARKER:" + new_manual_text
            st.toast(f"Manual changes for {selected_loc} staged!")

        # --- SECTION 3: SUBMIT AND BUILD ---
        st.divider()
        if st.button("🚀 Submit & Build Final Files"):
            # Combine Table Replacements and Manual Edits
            final_replacements = dict(zip(edited_df["Elaborated Location"], edited_df["Replacement Word"]))
            final_replacements.update(st.session_state.manual_edits)
            
            with st.spinner("Processing final output..."):
                # Standard Wrap
                standard_bytes, _ = process_zip_recursive(uploaded_file.getvalue(), search_word, final_replacements)
                
                # Pure XML Extraction
                flat_buffer = BytesIO()
                with zipfile.ZipFile(flat_buffer, 'w', zipfile.ZIP_DEFLATED) as flat_zip:
                    process_zip_recursive(uploaded_file.getvalue(), search_word, final_replacements, flatten_to_zip=flat_zip)
                
                st.success("Files Rebuilt! Standard compression may result in smaller file sizes.")
                
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(f"📥 Download Original Format", data=standard_bytes, file_name=f"modified_{uploaded_file.name}")
                with c2:
                    st.download_button(f"📥 Download Pure XML Extract (.zip)", data=flat_buffer.getvalue(), file_name="pure_xml_extract.zip")