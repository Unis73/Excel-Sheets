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
    # Check if the series contains only text and no numbers
    return series.apply(lambda x: isinstance(x, str) and not any(char.isdigit() for char in x)).all()

def main():
    st.title("Excel Data Management")

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
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Save the uploaded file to a temporary file path
        if 'original_file_path' not in st.session_state:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getbuffer())
                st.session_state.original_file_path = temp_file.name

        if 'df' not in st.session_state:
            df = load_data(st.session_state.original_file_path)
            df = clean_data(df)
            st.session_state.df = df
        else:
            df = st.session_state.df

        # Initialize form data if not present
        if 'form_data' not in st.session_state:
            st.session_state.form_data = {col: '' for col in df.columns}

        st.write('Current Data:')
        st.write(st.session_state.df)

        # Sidebar form fields
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

        # Sidebar buttons
        col1, col2 = st.sidebar.columns([2, 1])
        
        with col1:
            if st.button('Add Data'):
                new_data = {col: new_data[col] if new_data[col] != '' else 'NA' for col in df.columns}
                new_data_df = pd.DataFrame([new_data])
                
                # Check for duplicate entries in the first column
                first_col_name = df.columns[0]  # Assuming the first column should be unique
                if new_data_df[first_col_name].values[0] in df[first_col_name].values:
                    st.sidebar.warning(f'The value "{new_data[first_col_name]}" already exists in the "{first_col_name}" column.')
                else:
                    st.session_state.df = pd.concat([st.session_state.df, new_data_df], ignore_index=True)
                    st.session_state.df = clean_data(st.session_state.df)
                    st.sidebar.success('Data added successfully!')
                    # Clear form data
                    st.session_state.form_data = {col: '' for col in df.columns}
                    st.experimental_rerun()

        with col2:
            if st.button('Clear All'):
                # Clear the form fields without refreshing
                st.session_state.form_data = {col: '' for col in df.columns}
                for col in df.columns:
                    st.session_state[f"{col}_input"] = ''
                st.experimental_rerun()

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

        # Filter and display data
        st.header('Retrieve Data')
        filter_cols = st.multiselect('Select columns for filter:', options=df.columns)
        
        filter_values = {}
        for col in filter_cols:
            filter_values[col] = st.text_input(f'Enter value to filter {col}:')

        filtered_df = df.copy()
        for col, value in filter_values.items():
            if value:
                filtered_df = filtered_df[filtered_df[col].str.lower() == value.lower()]

        # Download filtered data
        if not filtered_df.empty:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
            buffer.seek(0)
            st.download_button(
                label="Download Filtered Data",
                data=buffer,
                file_name="filtered_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
