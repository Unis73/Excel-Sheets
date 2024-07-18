import streamlit as st
import pandas as pd
import openpyxl 

# Function to load Excel data
@st.cache_data  # Updated cache decorator
def load_data(file):
    return pd.read_excel(file)

# Function to save data back to Excel
def save_data(data, file):
    data.to_excel(file, index=False)

def main():
    st.title('Excel Data Entry and Retrieval')

    # Sidebar for file upload and data entry
    st.sidebar.title('Data Entry')
    uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Load data from uploaded file
        df = load_data(uploaded_file)

        # Show the current data in a table
        st.write('Current Data:')
        st.write(df)

        # Data entry form
        st.sidebar.header('Enter New Data')
        new_data = {}
        for col in df.columns:
            new_data[col] = st.sidebar.text_input(col)

        # Button to add new data
        if st.sidebar.button('Add Data'):
            df = df.append(new_data, ignore_index=True)
            save_data(df, uploaded_file)
            st.sidebar.success('Data added successfully!')

        # Filter and display data
        st.header('Retrieve Data')
        st.write('Filter data:')
        filter_col = st.selectbox('Select column for filter:', options=df.columns)
        filter_value = st.text_input(f'Enter value to filter {filter_col}:')
        filtered_df = df[df[filter_col] == filter_value]
        st.write(filtered_df)

if __name__ == '__main__':
    main()
