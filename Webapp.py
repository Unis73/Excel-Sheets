import streamlit as st
import pandas as pd
import openpyxl
import tempfile

# Function to load Excel data
@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    return df

# Function to save data back to Excel
def save_data(data):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        data.to_excel(tmp.name, index=False)
        return tmp.name

def clean_data(df):
    # Convert all columns to string type to handle mixed types
    df = df.astype(str)
    return df

def main():
    st.title("Excel Data Loader")

    # Sidebar for file upload and data entry
    st.sidebar.title('Data Entry')
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        df = clean_data(df)
        
        st.write("Filtered Data:")

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
            file_path = save_data(df)
            st.sidebar.success('Data added successfully!')
            st.sidebar.markdown(f"[Download updated file](file://{file_path})")

        # Filter and display data
        st.header('Retrieve Data')
        st.write('Filter data:')
        filter_col = st.selectbox('Select column for filter:', options=df.columns)
        filter_value = st.text_input(f'Enter value to filter {filter_col}:')

        # Filter the DataFrame
        if filter_value:
            filtered_df = df[df[filter_col] == filter_value]
            st.write(filtered_df)

if __name__ == "__main__":
    main()
