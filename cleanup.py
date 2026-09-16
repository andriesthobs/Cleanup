import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------

st.set_page_config(
    page_title="Shoprite Incident Cleanup",
    page_icon="🧹",
    layout="wide"
)

st.title("🧹 Shoprite Incident Cleanup")
st.markdown(
    """
    Upload an Excel file containing a **Cleaned** sheet.
    
    Matching incidents will automatically be moved to **Layer2** and a new workbook will be generated for download.
    """
)

# --------------------------------------------------
# SEARCH VALUES
# --------------------------------------------------

search_values = {
    "066363", "099946", "099823", "061444", "099718", "069620", "086385", "070386",
    "052413", "085795", "071277", "094899", "094873", "085614", "070174", "070166",
    "085494", "030504", "066583", "080486", "061509", "057405", "054392", "041456",

    "085622", "043937", "054237", "055267", "038722", "088280", "099962", "066305",
    "035839", "042525", "071722", "085630", "070190", "042842", "099166", "070556",
    "071633", "058011", "060977", "039532", "089804", "071675", "070051", "060773",

    "042779", "30986", "058752", "089511", "030978", "061127", "061054", "061020",
    "066800", "068747", "078075", "037394", "083426", "61151", "82763", "089846",
    "070988", "089676", "057463", "058265", "83832", "061575", "083670", "083044",

    "087064", "060341", "070077", "064133", "061101", "089838", "060668", "083989",
    "030520", "83824", "085672", "046529", "044030", "061517", "089870", "069646",
    "070506", "091493", "061533", "69670", "093322", "058590", "035805", "058689",
    "168024", "368771", "55453", "55330", "068797", "080745", "043945", "83175",
    "43741", "52439", "49959", "161674", "61258", "041375", "63242", "035994",
    "099920", "038170", "087917", "064921", "092758", "002379", "070328", "032629",
    "050461", "030376", "036225", "162735", "163113", "083832", "0720", "050348",
    "066143", "083230", "49933", "52421", "001608", "091249", "063218", "162858",
    "162052", "169842", "166357", "168977", "168480", "169054", "163139", "010516",
    "170451", "04791", "083531", "165220", "168406", "170605", "169973", "166161",
    "168969", "164101", "166658", "035960", "0252", "68755", "0623", "0729",
    "161991", "047779", "034817", "034045", "069654", "058281", "096906", "053257",
    "089773", "099857", "055657", "067628", "007840", "070904", "033984", "085127",
    "006802", "096605", "099988", "061232", "000107", "081408", "092156", "057188",
    "050550", "85494", "6399", "34089", "60668",

    # NEW VALUES
    "JWL",
    "SRGP"
}

# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Shoprite Incident File",
    type=["xlsx"]
)

# --------------------------------------------------
# PROCESS FILE
# --------------------------------------------------

if uploaded_file:

    try:

        df_cleaned = pd.read_excel(
            uploaded_file,
            sheet_name="Cleaned",
            dtype=str,
            engine="openpyxl"
        )

        col_c = df_cleaned.columns[2]

        def find_matches(value):
            value = str(value) if pd.notna(value) else ""

            matches = []

            for code in search_values:
                if code.upper() in value.upper():
                    matches.append(code)

            return matches

        match_list = df_cleaned[col_c].apply(find_matches)

        mask = match_list.apply(lambda x: len(x) > 0)

        layer2_df = df_cleaned[mask].copy()
        cleaned_df = df_cleaned[~mask].copy()

        # ------------------------------------------
        # METRICS
        # ------------------------------------------

        total_incidents = len(df_cleaned)
        moved_incidents = len(layer2_df)
        remaining_incidents = len(cleaned_df)

        percentage = (
            round((moved_incidents / total_incidents) * 100, 2)
            if total_incidents > 0
            else 0
        )

        st.success(
            f"✅ Cleanup Complete - {moved_incidents:,} incidents moved to Layer2"
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total Incidents", f"{total_incidents:,}")

        with col2:
            st.metric("Moved to Layer2", f"{moved_incidents:,}")

        with col3:
            st.metric("Remaining in Cleaned", f"{remaining_incidents:,}")

        # ------------------------------------------
        # SUMMARY
        # ------------------------------------------

        st.subheader("📊 Cleanup Summary")

        st.info(
            f"""
            • Total incidents analysed: **{total_incidents:,}**

            • Incidents moved to Layer2: **{moved_incidents:,}**

            • Incidents remaining in Cleaned: **{remaining_incidents:,}**

            • Percentage moved: **{percentage}%**
            """
        )

        # ------------------------------------------
        # BREAKDOWN
        # ------------------------------------------

        match_counts = {}

        for code in search_values:

            count = df_cleaned[col_c].fillna("").astype(str).str.upper().str.contains(
                code.upper(),
                na=False,
                regex=False
            ).sum()

            if count > 0:
                match_counts[code] = count

        if match_counts:

            breakdown_df = pd.DataFrame(
                list(match_counts.items()),
                columns=["Match Value", "Incidents"]
            )

            breakdown_df = breakdown_df.sort_values(
                by="Incidents",
                ascending=False
            )

            st.subheader("📈 Incident Breakdown")

            st.dataframe(
                breakdown_df,
                use_container_width=True,
                hide_index=True
            )

        # ------------------------------------------
        # PREVIEW
        # ------------------------------------------

        tab1, tab2 = st.tabs(
            ["Cleaned (Remaining)", "Layer2 (Moved)"]
        )

        with tab1:
            st.subheader(
                f"Remaining Records ({remaining_incidents:,})"
            )
            st.dataframe(
                cleaned_df.head(50),
                use_container_width=True
            )

        with tab2:
            st.subheader(
                f"Moved Records ({moved_incidents:,})"
            )
            st.dataframe(
                layer2_df.head(50),
                use_container_width=True
            )

        # ------------------------------------------
        # OUTPUT FILE
        # ------------------------------------------

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            cleaned_df.to_excel(
                writer,
                sheet_name="Cleaned",
                index=False
            )

            layer2_df.to_excel(
                writer,
                sheet_name="Layer2",
                index=False
            )

        output.seek(0)

        filename = (
            f"Shoprite_Incident_Cleanup_"
            f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        st.download_button(
            label="📥 Download Processed Workbook",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error processing file: {e}")