import streamlit as st
import pandas as pd
from adyen_batches import create_excel_with_formulas

mapping_data = [
    ["FCRotterdam","Store.FabienneChapot_POS_NL.005","De Meent 124","3011JS","Rotterdam","NL","5","1273","ROT"],
    ["FCRosada","Store.FabienneChapot_POS_NL.FCRosada","Rosada 113","4703TB","Roosendaal","NL","FCRosada","1275","ROS"],
    ["FCAmsterdam","Store.FabienneChapot_POS_NL.NL_001","Hartenstraat 7","1016 BZ","Amsterdam","NL","NL_001","1270","AMS"],
    ["FCHaarlem","Store.FabienneChapot_POS_NL.NL_002","Barteljorisstraat 24","2011 RB","Haarlem","NL","NL_002","1271","HRL"],
    ["FCMaastricht","Store.FabienneChapot_POS_NL.NL_003","Maastrichter Brugstraat 4","6211 ET","Maastricht","NL","NL_003","1272","MAA"],
    ["FCSampleSale","Store.FabienneChapot_POS_NL.NL_009","Herengracht  479","1017BS","Amsterdam","NL","NL_009","1276","SAM"],
    ["FCMall of the Netherlands","Store.FabienneChapot_POS_NL.NL_010","Liguster  202","2262AC","Leidschendam","NL","NL_010","1277","MOTN"],
    ["FCBreda","Store.FabienneChapot_POS_NL.NL_011","Ridderstraat 7","4811JA","Breda","NL","NL_011","1278","BRE"],
    ["FCRoermond","Store.FabienneChapot_POS_NL.NL_012","Stadsweide 212","6041TD","Roermond","NL","NL_012","1279","ROE"],
    ["FCDenBosch","Store.FabienneChapot_POS_NL.NL_013","Kerkstraat 5","5211KD","Den Bosch","NL","NL_013","1269","DBO"],
    ["FC MOTN Archive Store","Store.FabienneChapot_POS_NL.NL_014","Liguster 19","2262CL","Leidschendam","NL","NL_014","1268","MOTNAR"],
    ["FCBataviastad","Store.FabienneChapot_POS_NL.NL_04","Bataviaplein 64","8242 PN","Lelystad","NL","NL_04","1274","BAT"],
]
mapping_columns = ["Description","Store key","Street","Postal code","City","Country code","Store ID","GL","Cost Centre"]
mapping_df = pd.DataFrame(mapping_data, columns=mapping_columns)

# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------

st.title("Adyen Batches")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    filename = uploaded_file.name.replace("settlement_detail_report_", "")
    filename = filename[:31]
    st.success("File uploaded successfully!")

    if st.button("Generate Excel File"):
        excel_file = create_excel_with_formulas(
            filename=filename,
            df=df,
            mapping_df=mapping_df
        )

        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=f"{uploaded_file.name.split(".")[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        )
