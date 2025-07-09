import streamlit as st
import pandas as pd
import io

st.title("🔍 Excel Lookup: Single & Multiple Columns")

st.write("""
Upload two Excel files below.  
You can choose to run a lookup on a single column or multiple columns.
If found, writes the value(s); otherwise, writes "Missing".
""")

# Upload two Excel files
file1 = st.file_uploader("Upload First Excel (Base file)", type=["xlsx"])
file2 = st.file_uploader("Upload Second Excel (Reference file)", type=["xlsx"])

if file1 and file2:
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    st.subheader("Step 1️⃣: Choose Lookup Mode")
    mode = st.radio("Select Lookup Mode", ["Single Column", "Multiple Columns"])

    if mode == "Single Column":
        col1 = st.selectbox("Select Column from First Excel", df1.columns)
        col2 = st.selectbox("Select Column from Second Excel", df2.columns)

    else:  # Multiple
        cols1 = st.multiselect("Select Columns from First Excel", df1.columns)
        cols2 = st.multiselect("Select Columns from Second Excel", df2.columns)

        if len(cols1) != len(cols2):
            st.error("⚠️ Please select the same number of columns in both sheets.")
            st.stop()

    if st.button("Run Lookup"):
        result_df = df1.copy()

        if mode == "Single Column":
            result_df["Lookup_Result"] = result_df[col1].apply(
                lambda x: x if x in df2[col2].values else "Missing"
            )
        else:  # Multiple
            # Create key columns
            df1["__key__"] = df1[cols1].astype(str).agg("|".join, axis=1)
            df2["__key__"] = df2[cols2].astype(str).agg("|".join, axis=1)

            result_df["Lookup_Result"] = result_df["__key__"].apply(
                lambda x: x if x in df2["__key__"].values else "Missing"
            )

            # Drop helper column
            result_df.drop(columns=["__key__"], inplace=True)

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
            file_name="lookup_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload both Excel files to proceed.")
