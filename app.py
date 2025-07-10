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
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        text-align: center;
        color: #333;
        padding: 5px;
        font-size: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🔍 Excel VLOOKUP + Match Tool")

file1 = st.file_uploader("📄 Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("📄 Upload Second Excel (Lookup file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    col1 = st.selectbox("Match Column from First Excel", df1.columns, key="match1")
    col2 = st.selectbox("Match Column from Second Excel", df2.columns, key="match2")

    fetch_col = st.selectbox(
        "Value Column to Fetch from Second Excel (optional)",
        ["(None)"] + list(df2.columns), key="fetch"
    )

    if st.button("🚀 Run Lookup"):
        result_df = df1.copy()

        if fetch_col != "(None)":
            # VLOOKUP mode
            lookup_dict = df2.set_index(col2)[fetch_col].to_dict()
            result_df["Result"] = result_df[col1].apply(
                lambda x: lookup_dict.get(x, "Not Available")
            )
        else:
            # Just match check → write value if found, else Not Available
            lookup_set = set(df2[col2])
            result_df["Result"] = result_df[col1].apply(
                lambda x: x if x in lookup_set else "Not Available"
            )

        st.dataframe(result_df)

        def to_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = to_excel(result_df)

        st.download_button(
            label="📥 Download Result Excel",
            data=excel_data,
            file_name="lookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("👆 Please upload both Excel files to begin.")


# ✅ Footer with developer name and laptop logo
st.markdown(
    """
    <div class='footer'>
        💻 Developed by <b>ER. Ruchi Tiwari</b>
    </div>
    """,
    unsafe_allow_html=True,
)
