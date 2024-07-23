import streamlit as st
import pandas as pd
import openpyxl
import tempfile
import os

# Function to load Excel data
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    return df

# Function to save data back to Excel
def save_data(data, file_path):
    data.to_excel(file_path, index=False)

def clean_data(df):
    # Convert all columns to string type to handle mixed types
    df = df.astype(str)
    return df

def main():
    st.title("Excel Data Loader")

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
        if 'df' not in st.session_state:
            # Load and clean data only if it's not already loaded
            df = load_data(uploaded_file)
            df = clean_data(df)
            st.session_state.df = df
            # Save original file path in session state
            st.session_state.file_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
            with open(st.session_state.file_path, 'wb') as temp_file:
                temp_file.write(uploaded_file.getbuffer())
        else:
            df = st.session_state.df

        st.write("Filtered Data:")

        # Show the current data in a table
        st.write('Current Data:')
        data_placeholder = st.empty()
        data_placeholder.write(df)

        # Data entry form
        st.sidebar.header('Enter New Data')
        new_data = {}
        for col in df.columns:
            new_data[col] = st.sidebar.text_input(col, key=col)

        # Button to add new data
        if st.sidebar.button('Add Data'):
            new_data_df = pd.DataFrame([new_data])
            df = pd.concat([df, new_data_df], ignore_index=True)
            st.session_state.df = df

            # Save data back to the temporary file path
            save_data(df, st.session_state.file_path)
            st.sidebar.success('Data added successfully!')

            # Refresh the displayed DataFrame
            data_placeholder.write(df)

        # Filter and display data
        st.header('Retrieve Data')
        st.write('Filter data:')
        filter_col = st.selectbox('Select column for filter:', options=df.columns)
        filter_value = st.text_input(f'Enter value to filter {filter_col}:')

        # Filter the DataFrame
        if filter_value:
            filtered_df = df[df[filter_col] == filter_value]
            st.write(filtered_df)

        # Provide a download link for the updated file
        with open(st.session_state.file_path, 'rb') as f:
            st.download_button('Download updated file', f, file_name='updated_file.xlsx')

if __name__ == "__main__":
    main()
