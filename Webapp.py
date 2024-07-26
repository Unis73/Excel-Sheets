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
    st.title("Excel Data Management")

    # Sidebar for file upload and data entry
    st.sidebar.title('Data Entry')
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Save the uploaded file to a temporary file path
        if 'original_file_path' not in st.session_state:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                temp_file.write(uploaded_file.getbuffer())
                st.session_state.original_file_path = temp_file.name

        if 'df' not in st.session_state or st.session_state.uploaded_file != uploaded_file:
            df = load_data(st.session_state.original_file_path)
            df = clean_data(df)
            st.session_state.df = df
            st.session_state.uploaded_file = uploaded_file
            st.experimental_rerun()  # Refresh the page when a new file is uploaded
        else:
            df = st.session_state.df

        st.write('Current Data:')
        st.write(st.session_state.df)

        st.sidebar.header('Enter New Data')

        # Initialize new_data dictionary
        if 'new_data' not in st.session_state:
            st.session_state.new_data = {col: '' for col in df.columns}

        # Add a persistent warning message for the first column
        first_col = df.columns[0]
        st.sidebar.warning("Make sure the first column is unique.")

        # Data entry form
        new_data = {}
        for col in df.columns:
            key = f"{col}_input"
            if key not in st.session_state:
                st.session_state[key] = ""
            if is_pure_text_column(df[col]):
                unique_values = df[col].unique().tolist()
                new_data[col] = st.sidebar.selectbox(f"Select or enter {col}", options=[""] + unique_values, key=key)
            else:
                new_data[col] = st.sidebar.text_input(f"{col}", value=st.session_state[key], key=key)

        # Buttons 
        add_button = st.sidebar.button('Add Data')
        clear_button = st.sidebar.button('Clear All')

        if add_button:
            new_data_cleaned = {col: new_data[col] if new_data[col] != '' else 'NA' for col in df.columns}
            new_data_df = pd.DataFrame([new_data_cleaned])

            first_col_name = df.columns[0]
            if new_data_df[first_col_name].values[0] in df[first_col_name].values:
                st.warning(f'The value "{new_data_cleaned[first_col_name]}" already exists in the "{first_col_name}" column.')
            else:
                st.session_state.df = pd.concat([st.session_state.df, new_data_df], ignore_index=True)
                st.session_state.df = clean_data(st.session_state.df)
                st.sidebar.success('Data added successfully!')
                
                # Clear the form after adding data
                for col in df.columns:
                    key = f"{col}_input"
                    st.session_state[key] = ""

                st.experimental_rerun()

        if clear_button:
            for col in df.columns:
                key = f"{col}_input"
                st.session_state[key] = ""
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
            st.success('File downloaded successfully!')

        # Filter and display data
        st.header('Retrieve Data')
        filter_cols = st.multiselect('Select columns for filter:', options=df.columns)

        filter_values = {}
        for col in filter_cols:
            filter_values[col] = st.text_input(f'Enter value to filter {col}:')

        if st.button('Apply Filters'):
            filtered_df = df.copy()
            for col, value in filter_values.items():
                if value:
                    filtered_df = filtered_df[filtered_df[col].str.lower() == value.lower()]

            if filtered_df.empty:
                st.warning('No data matches the filter criteria.')
            else:
                # Display filtered data
                st.write('Filtered Data:')
                st.write(filtered_df)

                # Download filtered data
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
