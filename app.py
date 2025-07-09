import streamlit as st
import pandas as pd
import io

st.title("🔍 Excel VLOOKUP App")

st.write("""
Upload two Excel files below.  
You can select the column to lookup and the column to fetch from the second file.
""")

# Upload two Excel files
file1 = st.file_uploader("Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel (Lookup file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("Step 1️⃣: Select Lookup Key Columns")
    col1 = st.selectbox("Select Key Column from First Excel", df1.columns)
    col2 = st.selectbox("Select Key Column from Second Excel", df2.columns)

    st.subheader("Step 2️⃣: Select Value Column from Second Excel")
    value_col = st.selectbox("Select Value Column to fetch from Second Excel", df2.columns)

    if st.button("Run VLOOKUP"):
        result_df = df1.copy()
        lookup_dict = df2.set_index(col2)[value_col].to_dict()
        result_df[f"VLOOKUP_{value_col}"] = result_df[col1].map(lookup_dict)

        st.success("✅ VLOOKUP Completed!")
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
            file_name="vlookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload both Excel files to proceed.")
