import streamlit as st
import pandas as pd
import io

st.title("🔍 Excel VLOOKUP + Match App")

st.write("""
Upload two Excel files below.  
Select a column from each file to match.  
If a match is found, fetches the value from the second file; otherwise writes "Missing".
""")

# Upload two Excel files
file1 = st.file_uploader("Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel (Lookup file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("Step 1️⃣: Select Columns to Match & Fetch")
    col1 = st.selectbox("Select Match Column from First Excel", df1.columns)
    col2 = st.selectbox("Select Match Column from Second Excel", df2.columns)
    fetch_col = st.selectbox("Select Value Column to fetch from Second Excel", df2.columns)

    if st.button("Run VLOOKUP + Match"):
        result_df = df1.copy()

        # Create lookup dictionary
        lookup_dict = df2.set_index(col2)[fetch_col].to_dict()

        result_df["Lookup_Result"] = result_df[col1].apply(
            lambda x: lookup_dict.get(x, "Missing")
        )

        st.success("✅ Lookup Completed!")
        st.dataframe(result_df)

        # Function to convert dataframe to Excel binary
        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = to_excel(result_df)

        st.download_button(
            label="📥 Download Result Excel",
            data=excel_data,
            file_name="lookup_match_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload both Excel files to proceed.")
