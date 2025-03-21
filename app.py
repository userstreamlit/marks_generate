import streamlit as st
import pandas as pd
import random
import os

# Check if template file exists
template_file_path = "Template.xlsx"
if not os.path.exists(template_file_path):
    st.error(f"Template file '{template_file_path}' not found in the directory. Please ensure it exists.")
    st.stop()

def generate_marks(total):
    # Function to generate marks respecting the constraints
    max_attempts = 1000  # To avoid infinite loops
    attempts = 0
    
    while attempts < max_attempts:
        mark1 = random.randint(0, 1)  # 0 or 1
        mark2 = random.randint(0, 5)  # 0 to 5
        mark3 = random.randint(0, 3)  # 0 to 3
        mark4 = random.randint(0, 2)  # 0 to 2
        mark5 = random.randint(0, 1)  # 0 or 1
        
        # Check if sum matches total
        if mark1 + mark2 + mark3 + mark4 + mark5 == total:
            return [mark1, mark2, mark3, mark4, mark5]
        
        attempts += 1
    
    # If no valid combination is found
    return None

# Streamlit app
st.title("Marks Generator for Excel Template")

# Instructions for the user
st.subheader("Instructions")
st.markdown("""
1. **Download the Template**
2. **Fill the Template**
3. **Reupload the excel file**
""")

# Download button for the template
st.subheader("Step 1: Download the Template")
with open(template_file_path, "rb") as f:
    st.download_button(
        label="Download Template Excel File",
        data=f,
        file_name="template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# Excel file upload section
st.subheader("Step 2: Upload Your Filled Excel File")
uploaded_file = st.file_uploader("Choose your filled Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Read excel file
    df_input = pd.read_excel(uploaded_file)
    
    # Check if required columns exist
    required_columns = ['Mark1', 'Mark2', 'Mark3', 'Mark4', 'Mark5', 'Total']
    if not all(col in df_input.columns for col in required_columns):
        st.error("Excel file must contain columns: Mark1, Mark2, Mark3, Mark4, Mark5, Total")
    else:
        # Check for invalid totals (greater than 12 or less than 0) and report row numbers
        invalid_rows = []
        for idx, total in enumerate(df_input['Total'], start=1):
            if pd.isna(total):
                invalid_rows.append((idx, "missing"))
            elif total > 12 or total < 0:
                invalid_rows.append((idx, total))
        
        if invalid_rows:
            # Display error message with row numbers and invalid totals
            error_msg = "The following rows have invalid totals (must be between 0 and 12):\n"
            for row, total in invalid_rows:
                if total == "missing":
                    error_msg += f"Row {row}: Total is missing\n"
                else:
                    error_msg += f"Row {row}: Total = {total}\n"
            st.error(error_msg)
        else:
            # If all totals are valid, proceed with mark generation
            df_input['Mark1'] = None
            df_input['Mark2'] = None
            df_input['Mark3'] = None
            df_input['Mark4'] = None
            df_input['Mark5'] = None
            
            # Generate marks for each row based on Total
            results = []
            for total in df_input['Total']:
                marks = generate_marks(total)
                if marks:
                    results.append(marks + [total])
                else:
                    st.error(f"Cannot generate marks for total: {total}")
                    break
            
            if len(results) == len(df_input):
                # Create result dataframe
                result_df = pd.DataFrame(results, columns=['Mark1', 'Mark2', 'Mark3', 'Mark4', 'Mark5', 'Total'])
                st.subheader("Step 3: View and Download Results")
                st.table(result_df)
                
                # Download button for results
                excel_buffer = pd.ExcelWriter('generated_marks.xlsx', engine='openpyxl')
                result_df.to_excel(excel_buffer, index=False)
                excel_buffer.close()
                
                with open('generated_marks.xlsx', 'rb') as f:
                    st.download_button(
                        label="Download Generated Marks Excel File",
                        data=f,
                        file_name="generated_marks.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )