import streamlit as st
import pandas as pd
import io

st.title("🔍 Excel Lookup with Value or Missing")

st.write("""
Upload two Excel files below.  
This tool checks if the values in one column of the first Excel exist in a column of the second Excel.  
If found, writes the value; otherwise, writes "Missing".
""")

# Upload two Excel files
file1 = st.file_uploader("Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel (Reference file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("Step 1️⃣: Select Columns to Compare")
    col1 = st.selectbox("Select Column from First Excel", df1.columns)
    col2 = st.selectbox("Select Column from Second Excel", df2.columns)

    if st.button("Run Lookup"):
        result_df = df1.copy()

        result_df["Lookup_Result"] = result_df[col1].apply(
            lambda x: x if x in df2[col2].values else "Missing"
        )

        st.success("✅ Lookup Completed!")
        st.dataframe(result_df)

        # Function to convert dataframe to Excel binary
        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            processed_data = output.getvalue()
            return processed_data

        excel_data = to_excel(result_df)

        st.download_button(
            label="📥 Download Result Excel",
            data=excel_data,
            file_name="lookup_with_value_or_missing.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload both Excel files to proceed.")
