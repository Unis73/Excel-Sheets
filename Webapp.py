import streamlit as st
import pandas as pd
import openpyxl
import tempfile
import os

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
    st.title("Excel Data")

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
        if 'file_path' not in st.session_state:
            # Save uploaded file to a temporary file path
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getbuffer())
                st.session_state.file_path = temp_file.name
                st.sidebar.write(f"File path: {st.session_state.file_path}")

        df = load_data(st.session_state.file_path)
        df = clean_data(df)
        st.session_state.df = df

        # Show the current data in a table
        st.write('Current Data:')
        data_placeholder = st.empty()
        data_placeholder.write(df)

        # Data entry form
        st.sidebar.header('Enter New Data')
        new_data = {}
        for col in df.columns:
            if is_pure_text_column(df[col]):
                # If column contains only text values with no numbers, use dropdown
                unique_values = df[col].unique().tolist()
                new_data[col] = st.sidebar.selectbox(
                    f"Select or enter {col}",
                    options=[""] + unique_values,
                    key=f"{col}_dropdown"
                )
            else:
                # If column contains numbers or mixed data, use text input
                new_data[col] = st.sidebar.text_input(f"{col}", key=f"{col}_input")

        # Button to add new data
        if st.sidebar.button('Add Data'):
            new_data = {col: new_data[col] if new_data[col] != '' else 'NA' for col in df.columns}
            new_data_df = pd.DataFrame([new_data])
            df = pd.concat([df, new_data_df], ignore_index=True)
            df = clean_data(df)
            st.session_state.df = df

            # Save data back to the original file path
            save_data(df, st.session_state.file_path)
            st.sidebar.success('Data added successfully!')

            # Refresh the displayed DataFrame
            data_placeholder.write(df)

        # Filter and display data
        st.header('Retrieve Data')
        filter_cols = st.multiselect('Select columns for filter:', options=df.columns)
        
        filter_values = {}
        for col in filter_cols:
            filter_values[col] = st.text_input(f'Enter value to filter {col}:')

        # Filter the DataFrame based on multiple conditions
        if filter_values:
            filtered_df = df.copy()
            for col, value in filter_values.items():
                if value:
                    filtered_df = filtered_df[filtered_df[col].str.lower() == value.lower()]
            st.write(filtered_df)

if __name__ == "__main__":
    main()
