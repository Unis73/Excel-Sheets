import streamlit as st
import pandas as pd
import openpyxl
import tempfile
import io

# Function to load Excel data
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to save data back to Excel
def save_data(data, file_path):
    data.to_excel(file_path, index=False)

# Function to clean data
def clean_data(df):
    df = df.fillna('NA').astype(str)
    return df

def is_pure_text_column(series):
    return series.apply(lambda x: isinstance(x, str) and not any(char.isdigit() for char in x)).all()

def main():
    st.title("EXCEL FORMS")

    # Hide specific Streamlit style elements
    hide_streamlit_style = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .css-18ni7ap.e8zbici2 {visibility: hidden;} /* Hide the Streamlit menu icon */
        .css-1v0mbdj.e8zbici1 {visibility: visible;} /* Keep the settings icon */
        </style>
    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)

    # Sidebar for file upload and data entry
    st.sidebar.title('Data Entry')
    st.sidebar.warning("Ensure the first column contains unique values.")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Check if a new file is uploaded
        if 'uploaded_file' in st.session_state and st.session_state.uploaded_file != uploaded_file:
            # Clear session state to reset the app if a new file is uploaded
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.experimental_rerun()

        if 'original_file_path' not in st.session_state:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getbuffer())
                st.session_state.original_file_path = temp_file.name
                st.session_state.uploaded_file = uploaded_file

        if 'df' not in st.session_state:
            df = load_data(st.session_state.original_file_path)
            df = clean_data(df)
            st.session_state.df = df
        else:
            df = st.session_state.df

        # Initialize form data if not present
        if 'form_data' not in st.session_state:
            st.session_state.form_data = {col: '' for col in df.columns}

        # Sidebar form fields for adding new data
        st.sidebar.header('Enter New Data')

        new_data = {}
        for col in df.columns:
            key = f"{col}_input"
            if is_pure_text_column(df[col]):
                unique_values = df[col].unique().tolist()
                new_data[col] = st.sidebar.selectbox(
                    f"Select or enter {col}",
                    options=[""] + unique_values,
                    key=key,
                    index=unique_values.index(st.session_state.form_data.get(col, '')) if st.session_state.form_data.get(col, '') in unique_values else 0
                )
            else:
                new_data[col] = st.sidebar.text_input(
                    f"{col}",
                    key=key,
                    value=st.session_state.form_data.get(col, '')
                )

        # Sidebar button to add data
        if st.sidebar.button('Add Data'):
            new_data = {col: new_data[col] if new_data[col] != '' else 'NA' for col in df.columns}
            new_data_df = pd.DataFrame([new_data])
            
            # Check for duplicate entries in the first column
            first_col_name = df.columns[0]  # Assuming the first column should be unique
            if new_data_df[first_col_name].values[0] in df[first_col_name].values:
                st.error(f'The value "{new_data[first_col_name]}" already exists in the "{first_col_name}" column.')
            else:
                st.session_state.df = pd.concat([st.session_state.df, new_data_df], ignore_index=True)
                st.session_state.df = clean_data(st.session_state.df)
                st.success('Data added successfully!')
                # Clear the form fields after successful data addition
                for col in df.columns:
                    key = f"{col}_input"
                    if key in st.session_state:
                        del st.session_state[key]
                st.experimental_rerun()  # Refresh the sidebar and form fields

        # Display and edit existing data
        st.header('Edit Data')
        edit_index = st.number_input('Select row to edit (index):', min_value=0, max_value=len(df)-1, step=1)
        
        if len(df) > 0:
            row_to_edit = df.iloc[edit_index]
            st.write(f'Editing row {edit_index}')
            
            edited_data = {}
            for col in df.columns:
                if is_pure_text_column(df[col]):
                    edited_data[col] = st.text_input(f"{col}", value=row_to_edit[col])
                else:
                    edited_data[col] = st.text_input(f"{col}", value=row_to_edit[col])
            
            if st.button('Save Changes'):
                df.iloc[edit_index] = edited_data
                st.session_state.df = clean_data(df)
                save_data(st.session_state.df, st.session_state.original_file_path)
                st.success('Changes saved successfully!')

        # Filter and display data
        st.header('Retrieve Data')
        filter_cols = st.multiselect('Select columns for filter:', options=df.columns)
        
        filter_values = {}
        filtered_df = df.copy()  # Initialize filtered_df before the condition

        if filter_cols:
            for col in filter_cols:
                filter_value = st.text_input(f'Enter value to filter {col}:')
                if filter_value:
                    filter_values[col] = filter_value
                    filtered_df = filtered_df[filtered_df[col].str.lower() == filter_value.lower()]

            if filtered_df.empty:
                st.warning('No matching records found.')
            else:
                st.write('Filtered Data:')
                st.write(filtered_df)

                # Download filtered data
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
                buffer.seek(0)
                
                if st.download_button(
                    label="Download Filtered Data",
                    data=buffer,
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ):
                    st.success('Download successfull!')

        # Create a download link for the updated data
        if st.button('Download Updated Data'):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as updated_file:
                save_data(st.session_state.df, updated_file.name)
                with open(updated_file.name, "rb") as file:
                    st.download_button(
                        label="Download Excel file",
                        data=file,
                        file_name="updated_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            st.success('Data downloaded successfully!')

if __name__ == "__main__":
    main()
