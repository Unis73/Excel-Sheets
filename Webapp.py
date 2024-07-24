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

        # Show the current data in a table
        st.write('Current Data:')
        edited_df = st.experimental_data_editor(df, key="editor")
        st.session_state.df = edited_df

        # Data entry form
        st.sidebar.header('Enter New Data')
        new_data = {}
        for col in df.columns:
            if is_pure_text_column(df[col]):
                unique_values = df[col].unique().tolist()
                new_data[col] = st.sidebar.selectbox(
                    f"Select or enter {col}",
                    options=[""] + unique_values,
                    key=f"{col}_dropdown"
                )
            else:
                new_data[col] = st.sidebar.text_input(f"{col}", key=f"{col}_input")

        # Button to add new data
        if st.sidebar.button('Add Data'):
            new_data = {col: new_data[col] if new_data[col] != '' else 'NA' for col in df.columns}
            new_data_df = pd.DataFrame([new_data])
            st.session_state.df = pd.concat([st.session_state.df, new_data_df], ignore_index=True)
            st.session_state.df = clean_data(st.session_state.df)
            st.sidebar.success('Data added successfully!')
            st.experimental_rerun()

        # Create a download link for the updated data
        if st.button('Download Updated Data'):
            updated_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
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

        if filter_values:
            filtered_df = df.copy()
            for col, value in filter_values.items():
                if value:
                    filtered_df = filtered_df[filtered_df[col].str.lower() == value.lower()]
            st.write(filtered_df)

if __name__ == "__main__":
    main()
