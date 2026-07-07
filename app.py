import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Universal Excel Merger", layout="centered")
st.title("📊 Universal Excel Merger")

# -- 1. KEY FILE --
st.subheader("🔑 Step 1: Upload Key File (Optional)")
key_file = st.file_uploader("Upload Key File (.xlsx)", type=["xlsx"], key="key_uploader")

column_mapping = {}
if key_file:
    try:
        key_df = pd.read_excel(key_file)
        for standard_name in key_df.columns:
            odd_names = key_df[standard_name].dropna().astype(str).tolist()
            for odd_name in odd_names:
                if odd_name.strip():
                    column_mapping[odd_name.strip()] = str(standard_name).strip()
        st.success("✅ Key file loaded!")
    except Exception as e:
        st.error(f"Error reading Key File: {e}")

# -- 2. UPLOAD DATA --
st.divider()
st.subheader("📂 Step 2: Upload Data Files")
uploaded_files = st.file_uploader("Upload your data files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# -- 3. CONDITIONAL SHEET SELECTION (MULTI-RULE) --
st.divider()
st.subheader("⚙️ Step 3: Conditional Sheet Rules (Optional)")
st.write("If certain files need specific sheets, add your rules below! All other files will just use their 1st sheet.")

# Create an empty table for the user to fill out
if "sheet_rules_df" not in st.session_state:
    st.session_state.sheet_rules_df = pd.DataFrame(
        [["", ""], ["", ""]], 
        columns=["If file name contains...", "...grab data from this exact sheet"]
    )

# Show the editable table
edited_rules_df = st.data_editor(st.session_state.sheet_rules_df, num_rows="dynamic", use_container_width=True)

# Convert the user's table into a clean Python list of rules
sheet_rules = []
for index, row in edited_rules_df.iterrows():
    kw = str(row["If file name contains..."]).strip()
    ts = str(row["...grab data from this exact sheet"]).strip()
    
    # Only add the rule if both boxes are filled out
    if kw and ts and kw != "nan" and ts != "nan":
        sheet_rules.append((kw.lower(), ts))

# -- MERGE EXECUTION --
if uploaded_files:
    if st.button("Merge Files", type="primary"):
        with st.spinner("Processing..."):
            all_data = []
            
            known_cols = {str(c).strip().lower() for c in set(column_mapping.keys()).union(set(column_mapping.values()))}
            
            with st.expander("🛠️ Processing Log", expanded=True):
                for file in uploaded_files: 
                    try:
                        st.write(f"⏳ **Processing:** `{file.name}`...")
                        
                        file_buffer = io.BytesIO(file.getvalue())
                        xls = pd.ExcelFile(file_buffer)
                        
                        sheet_to_use = 0  # Default: pretend no rules exist
                        target_sheet_clean = None
                        
                        # 💡 THE FIX: Check this file against EVERY rule in the list
                        for keyword, target in sheet_rules:
                            if keyword in file.name.lower():
                                target_sheet_clean = target
                                st.write(f"  ↳ 🎯 Rule matched ('{keyword}')! Looking for sheet `{target_sheet_clean}`...")
                                break # Stop checking rules once we find the first match
                        
                        if target_sheet_clean:
                            target_lower = target_sheet_clean.lower()
                            found_match = False
                            
                            for actual_sheet in xls.sheet_names:
                                if actual_sheet.strip().lower() == target_lower:
                                    sheet_to_use = actual_sheet
                                    found_match = True
                                    break
                            
                            if not found_match:
                                st.warning(f"  ↳ ⚠️ Sheet '{target_sheet_clean}' not found! Falling back to 1st sheet.")
                        else:
                            st.write(f"  ↳ ⏭️ No rules matched. Defaulting to 1st sheet.")
                        
                        # Auto-detect headers on the correct sheet
                        temp_df = pd.read_excel(xls, sheet_name=sheet_to_use, header=None, nrows=20)
                        
                        best_row_idx = 0
                        if known_cols:
                            max_matches = 0
                            for idx, row in temp_df.iterrows():
                                matches = sum(1 for cell in row if str(cell).strip().lower() in known_cols)
                                if matches > max_matches:
                                    max_matches = matches
                                    best_row_idx = idx
                        
                        # Final read
                        df = pd.read_excel(xls, sheet_name=sheet_to_use, header=best_row_idx)
                        
                        if column_mapping:
                            df = df.rename(columns=column_mapping)
                        
                        all_data.append(df)
                        st.write(f"✅ **Success:** Grabbed sheet `{sheet_to_use}`! ({len(df)} rows)")
                        
                    except Exception as e:
                        st.write(f"❌ **FAILED on {file.name}:** {e}")
            
            # -- COMBINE AND DOWNLOAD --
            if all_data:
                merged_df = pd.concat(all_data, ignore_index=True)
                st.success(f"🎉 Successfully merged ALL {len(all_data)} file(s)!")
                
                st.write("### Data Preview")
                st.dataframe(merged_df.head(10))
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    merged_df.to_excel(writer, index=False, sheet_name='Merged_Data')
                output.seek(0)
                
                st.download_button("⬇️ Download Merged Excel File", data=output, file_name="merged_output.xlsx")