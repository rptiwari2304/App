import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel VLOOKUP + Match Tool", layout="centered")

st.markdown(
    """
    <style>
    .main {background-color: #f8f9fa;}
    .stButton>button {background-color: #4CAF50; color:white;}
    .stDownloadButton>button {background-color: #007BFF; color:white;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🔍 Excel VLOOKUP + Match Tool")

st.write("""
Upload two Excel files below.  
This tool works in two ways:
- ✅ If you select only match columns → shows `Found` / `Missing`
- ✅ If you also select a value column → does a full VLOOKUP
""")

# Upload two Excel files
file1 = st.file_uploader("📄 Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("📄 Upload Second Excel (Lookup file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("Step 1️⃣: Select Columns to Match")
    col1 = st.selectbox("🔷 Match Column from First Excel", df1.columns, key="match1")
    col2 = st.selectbox("🔷 Match Column from Second Excel", df2.columns, key="match2")

    st.subheader("Step 2️⃣: (Optional) Select Value Column to Fetch")
    fetch_col = st.selectbox(
        "📌 Value Column from Second Excel (leave blank for Found/Missing)",
        ["(None)"] + list(df2.columns), key="fetch"
    )

    if st.button("🚀 Run Lookup"):
        result_df = df1.copy()

        if fetch_col != "(None)":
            # VLOOKUP mode
            lookup_dict = df2.set_index(col2)[fetch_col].to_dict()
            result_df["Result"] = result_df[col1].apply(
                lambda x: lookup_dict.get(x, "Missing")
            )
            st.success("✅ VLOOKUP Completed!")
        else:
            # Only match check
            lookup_set = set(df2[col2])
            result_df["Result"] = result_df[col1].apply(
                lambda x: "Found" if x in lookup_set else "Missing"
            )
            st.success("✅ Match Check Completed!")

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
    st.info("👆 Please upload both Excel files to begin.")
