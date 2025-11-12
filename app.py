import streamlit as st
import pandas as pd
from io import BytesIO
import re
import chardet

st.set_page_config(page_title="Text to Excel Converter", page_icon="📊", layout="wide")

st.title("📊 RCPier Report to Excel Converter")
st.markdown("Upload a text file containing RCPier load case data and convert it to an Excel file with organized sheets.")

# File upload
uploaded_file = st.file_uploader("Choose a text file", type=['txt'])

if uploaded_file is not None:
    # Read the file content
    file_bytes = uploaded_file.read()
    
    # Detect encoding
    result = chardet.detect(file_bytes)
    encoding = result['encoding']
    
    # Decode the file content
    try:
        textog = file_bytes.decode(encoding)
    except:
        # Fallback to utf-8 if detected encoding fails
        textog = file_bytes.decode('utf-8', errors='ignore')
    
    # Normalize line endings to Unix to make pattern matching robust
    textog = textog.replace('\r\n', '\n').replace('\r', '\n')
    
    # Process the text
    # Remove everything after "Selected load groups"
    if "Selected load groups" in textog:
        textog = textog[:textog.find("Selected load groups")]
    
    # Function to convert load case to dataframe
    def convSPtoDF(x, y):
        lines = [re.split(r'\s{2,}', line.strip()) for line in x.strip().split('\n')]
        df = pd.DataFrame(lines)
        return df
    
    # Extract tables from text file
    text = textog
    startp = "\n         -------------------------------------------------\n"
    endp = "\n \n      Auto generation details"
    sep_line = "-------------------------------------------------"
    
    dframedc = pd.DataFrame()
    dframell = pd.DataFrame()
    dframebr = pd.DataFrame()
    dframews = pd.DataFrame()
    dframewl = pd.DataFrame()
    
    df_dict = {}
    loadnameindex = text.find("Loadcase ID:")
    initial_loadnameindex = loadnameindex  # Save for debugging
    i = 1
    max_iterations = 1000  # Safety limit to prevent infinite loops
    
    # Debug counters
    skipped_empty_data = 0
    skipped_wrong_columns = 0
    skipped_empty_dataframe = 0
    processed_count = 0
    skipped_pattern_not_found = 0
    sample_extracted_data = None
    sample_df_shape = None
    
    # Create progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    with st.spinner("Processing load cases..."):
        while loadnameindex != -1 and i <= max_iterations:
            # Update progress
            status_text.text(f"Processing load case {i}...")
            progress_bar.progress(min(i / max_iterations, 1.0))
            # Extract load case name
            if loadnameindex + 13 >= len(text):
                break
            if text[loadnameindex+13] != "W":
                name = text[loadnameindex+13:loadnameindex + 17]
            else:
                name = text[loadnameindex+13:loadnameindex + 45]
                name = name.replace("    Name: ", "-")
                name = name.replace("\n", "")
            # Normalize whitespace in name to avoid trailing spaces
            name = name.strip()
            
            # Anchor to the 'Bearing loads:' section for consistent parsing
            bearing_hdr_idx = text.find("Bearing loads:", loadnameindex)
            if bearing_hdr_idx == -1:
                skipped_pattern_not_found += 1
                # Try to find next load case
                next_load = text.find("Loadcase ID:", loadnameindex + 1)
                if next_load == -1:
                    break
                text = text[next_load:]
                loadnameindex = 0
                i += 1
                continue
            
            # Find the dashed separator line after the 'Bearing loads:' header
            sep_idx = text.find(sep_line, bearing_hdr_idx)
            if sep_idx == -1:
                skipped_pattern_not_found += 1
                next_load = text.find("Loadcase ID:", loadnameindex + 1)
                if next_load == -1:
                    break
                text = text[next_load:]
                loadnameindex = 0
                i += 1
                continue
            
            # Data starts after the newline that follows the dashed line
            newline_after_sep = text.find('\n', sep_idx)
            if newline_after_sep == -1:
                data_start = sep_idx + len(sep_line)
            else:
                data_start = newline_after_sep + 1
            
            # Data ends at the next load case or at the auto-generation details marker, whichever comes first
            next_load = text.find("Loadcase ID:", loadnameindex + 1)
            auto_end = text.find("Auto generation details", data_start)
            end_candidates = [idx for idx in [next_load, auto_end] if idx != -1]
            if end_candidates:
                end_idx = min(end_candidates)
            else:
                end_idx = len(text)
            
            # Convert to dataframe
            data = text[data_start:end_idx]
            
            # Save sample data from first iteration for debugging
            if i == 1 and sample_extracted_data is None:
                sample_extracted_data = data[:500] if len(data) > 500 else data
            
            if not data.strip():  # Skip if no data
                skipped_empty_data += 1
                text = text[end_idx:]
                loadnameindex = text.find("Loadcase ID:")
                i += 1
                continue
            
            df = convSPtoDF(data, name)
            
            # Save sample DataFrame shape from first iteration for debugging
            if i == 1 and sample_df_shape is None:
                sample_df_shape = df.shape
            
            # Keep only the first 4 columns (delete any extra columns)
            if df.shape[1] > 4:
                df = df.iloc[:, :4]
            elif df.shape[1] < 4:
                # If less than 4 columns, skip this load case
                skipped_wrong_columns += 1
                text = text[end_idx:]
                loadnameindex = text.find("Loadcase ID:")
                i += 1
                continue
            
            # Add header
            if df.shape[1] == 4 and len(df) > 0:  # Also check if dataframe has rows
                processed_count += 1
                # Ensure unique column names by prefixing with load case name
                df.columns = [f"{name} - Line#", f"{name} - Bearing#", f"{name} - Direction", f"{name} - Loads-Kips"]
                
                # Insert a column with the load case identifier (unique per case)
                df.insert(0, name, [None] * len(df))
                
                # Store in dictionary
                df_dict[name] = df
                
                # Categorize by load type
                if "DC" in name:
                    dframedc = pd.concat([dframedc, df], axis=1)
                elif "WS" in name:
                    dframews = pd.concat([dframews, df], axis=1)
                elif "BR" in name:
                    dframebr = pd.concat([dframebr, df], axis=1)
                elif "WL" in name:
                    dframewl = pd.concat([dframewl, df], axis=1)
                else:
                    dframell = pd.concat([dframell, df], axis=1)
            else:
                # DataFrame has 4 columns but no rows
                skipped_empty_dataframe += 1
            
            # Move to next load case - ensure we advance past current position
            next_pos = end_idx
            if next_pos >= len(text):
                break
            text = text[next_pos:]
            loadnameindex = text.find("Loadcase ID:")
            i += 1
    
    # Clear progress indicators
    progress_bar.empty()
    status_text.empty()
    
    if i > max_iterations:
        st.warning(f"⚠️ Processing stopped after {max_iterations} iterations. Some load cases may not have been processed.")
    
    # Display summary
    if len(df_dict) > 0:
        st.success(f"✅ Processed {len(df_dict)} load cases!")
    else:
        st.error("❌ No load cases were processed. Please check the file format.")
        st.info("💡 Make sure the file contains 'Loadcase ID:' markers and the expected data format.")
        
        # Debug information
        with st.expander("🔍 Debug Information", expanded=True):
            st.write("**File Analysis:**")
            
            # Check for key patterns
            has_loadcase_id = "Loadcase ID:" in textog
            has_start_pattern = startp.strip() in textog
            has_end_pattern = endp.strip() in textog
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"Found 'Loadcase ID:': {'✅ Yes' if has_loadcase_id else '❌ No'}")
            with col2:
                st.write(f"Found start pattern: {'✅ Yes' if has_start_pattern else '❌ No'}")
            with col3:
                st.write(f"Found end pattern: {'✅ Yes' if has_end_pattern else '❌ No'}")
            
            # Count occurrences
            loadcase_count = textog.count("Loadcase ID:")
            st.write(f"**Number of 'Loadcase ID:' found:** {loadcase_count}")
            st.write(f"**Processing iterations completed:** {i-1}")
            st.write(f"**Successfully processed:** {processed_count}")
            st.write(f"**Skipped (pattern not found):** {skipped_pattern_not_found}")
            st.write(f"**Skipped (empty data):** {skipped_empty_data}")
            st.write(f"**Skipped (wrong column count):** {skipped_wrong_columns}")
            st.write(f"**Skipped (empty DataFrame):** {skipped_empty_dataframe}")
            
            # Show sample extracted data from first iteration
            if sample_extracted_data is not None:
                st.write("**Sample extracted data from first iteration:**")
                st.code(sample_extracted_data, language='text')
                if sample_df_shape is not None:
                    st.write(f"**First iteration DataFrame shape:** {sample_df_shape} (rows x columns)")
            
            # Check if loop exited early
            if initial_loadnameindex == -1:
                st.warning("⚠️ No 'Loadcase ID:' was found in the file, so the loop never started")
            elif i > max_iterations:
                st.warning(f"⚠️ Processing stopped at maximum iterations ({max_iterations})")
            elif loadnameindex == -1:
                st.info("ℹ️ Loop exited because no more 'Loadcase ID:' patterns were found")
            
            # Show sample text around first Loadcase ID if found
            if has_loadcase_id:
                first_idx = textog.find("Loadcase ID:")
                sample_start = max(0, first_idx - 100)
                sample_end = min(len(textog), first_idx + 500)
                st.write("**Sample text around first 'Loadcase ID:':**")
                st.code(textog[sample_start:sample_end], language='text')
                
                # Try to extract and show what data would be extracted
                try:
                    test_start = textog.find(startp, first_idx)
                    test_end = textog.find(endp, first_idx)
                    if test_start != -1 and test_end != -1 and test_end > test_start:
                        extracted_data = textog[test_start+len(startp):test_end]
                        st.write("**Sample extracted data (between patterns):**")
                        if extracted_data.strip():
                            st.code(extracted_data[:500] + ("..." if len(extracted_data) > 500 else ""), language='text')
                            # Try to create a test DataFrame
                            test_df = convSPtoDF(extracted_data, "TEST")
                            st.write(f"**Test DataFrame shape:** {test_df.shape} (rows x columns)")
                            if len(test_df) > 0:
                                st.write("**First few rows of test DataFrame:**")
                                st.dataframe(test_df.head(), use_container_width=True)
                        else:
                            st.warning("⚠️ Extracted data is empty!")
                except Exception as e:
                    st.write(f"**Error testing extraction:** {str(e)}")
            
            # Show the patterns being searched for
            st.write("**Patterns being searched:**")
            st.write(f"- Start pattern: `{repr(startp)}`")
            st.write(f"- End pattern: `{repr(endp)}`")
            
            # Try to find the actual pattern in the file
            if has_loadcase_id:
                first_idx = textog.find("Loadcase ID:")
                # Look for the separator line near the first load case
                search_area = textog[first_idx:first_idx+1000]
                if "-------------------------------------------------" in search_area:
                    sep_idx = search_area.find("-------------------------------------------------")
                    context_start = max(0, sep_idx - 20)
                    context_end = min(len(search_area), sep_idx + 50)
                    actual_pattern = search_area[context_start:context_end]
                    st.write("**Actual separator pattern found in file:**")
                    st.code(repr(actual_pattern), language='text')
    
    # Show statistics
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("DC Cases", len([k for k in df_dict.keys() if "DC" in k]))
    with col2:
        st.metric("LL Cases", len([k for k in df_dict.keys() if "LL" in k and "DC" not in k and "WS" not in k and "BR" not in k and "WL" not in k]))
    with col3:
        st.metric("BR Cases", len([k for k in df_dict.keys() if "BR" in k]))
    with col4:
        st.metric("WS Cases", len([k for k in df_dict.keys() if "WS" in k]))
    with col5:
        st.metric("WL Cases", len([k for k in df_dict.keys() if "WL" in k]))
    
    # Create Excel file in memory only if there's data to write
    has_data = (not dframedc.empty or not dframell.empty or not dframebr.empty or 
                not dframews.empty or not dframewl.empty)
    
    if has_data:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if not dframedc.empty:
                dframedc.to_excel(writer, sheet_name='DC', index=False)
            if not dframell.empty:
                dframell.to_excel(writer, sheet_name='LL', index=False)
            if not dframebr.empty:
                dframebr.to_excel(writer, sheet_name='BR', index=False)
            if not dframews.empty:
                dframews.to_excel(writer, sheet_name='WS', index=False)
            if not dframewl.empty:
                dframewl.to_excel(writer, sheet_name='WL', index=False)
        
        output.seek(0)
        
        # Download button
        excel_filename = uploaded_file.name.replace(".txt", ".xlsx")
        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name=excel_filename,
            mime="application/vnd.openpyxl-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No data to export. Excel file will not be created.")
    
    # Display preview of data
    st.subheader("📋 Data Preview")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["DC", "LL", "BR", "WS", "WL"])
    
    with tab1:
        if not dframedc.empty:
            st.dataframe(dframedc, use_container_width=True)
        else:
            st.info("No DC load cases found.")
    
    with tab2:
        if not dframell.empty:
            st.dataframe(dframell, use_container_width=True)
        else:
            st.info("No LL load cases found.")
    
    with tab3:
        if not dframebr.empty:
            st.dataframe(dframebr, use_container_width=True)
        else:
            st.info("No BR load cases found.")
    
    with tab4:
        if not dframews.empty:
            st.dataframe(dframews, use_container_width=True)
        else:
            st.info("No WS load cases found.")
    
    with tab5:
        if not dframewl.empty:
            st.dataframe(dframewl, use_container_width=True)
        else:
            st.info("No WL load cases found.")
    
    # Show all load case names
    with st.expander("📝 View all load case names"):
        st.write(list(df_dict.keys()))

else:
    st.info("👆 Please upload a text file to get started.")
    st.markdown("""
    ### How to use:
    1. Upload a text file containing RCPier load case data
    2. The app will automatically process the file
    3. Download the converted Excel file with organized sheets
    4. Preview the data in the tabs below
    """)
