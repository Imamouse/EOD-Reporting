import streamlit as st #pip install streamlit
import pandas as pd #pip install pandas
import plotly.express as px #pip install plotly-express
import numpy as np #pip install
import base64 #Standard python module
from io import StringIO, BytesIO # Standard python module
import time

def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "rts_rto" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)




st.set_page_config(page_title="Excel Plotter", layout="wide")
st.title("Excel Plotter 🗃" )
st.subheader("Drag me an excel file.")

hide_menu_style = """
        <style>
        #MainMenu {visibility: hidden;}
	    footer {visibility: hidden;}
        </style>
        """
st.markdown(hide_menu_style, unsafe_allow_html=True)

uploaded_file = st.file_uploader("Choose a xslx file", )
if uploaded_file:
    st.markdown("----")
    df = pd.read_excel(uploaded_file, sheet_name="REFUND_RTS.RTO ", 
 engine="openpyxl")
    

st.sidebar.header("Please Filter Here:")
status = st.sidebar.multiselect(
    "Select the Status:",
    options=df["Status"].unique(),
    default=df["Status"].unique(),
)

df_selection = df.query (
    "Status == @status"
)

df_selected = df_selection.drop(columns=["Entry Date",
                 "Handler",
                 "Date ticket was handled by Inspiro Esca Team (Date 1st Escalated)2",
                 "Email",
                 "Delivery Address",
                 "Date of order transaction",
                 "Remarks (To be filled by E-channel Team)",
                 "ESCA Comment", 
                 "Date refunded (To be filled by E-channel Team)"               
])

 
df_selected = df_selected.style.format({"Delivery no.": lambda x : '{0}'.format(x)})
 
st.dataframe(df_selected)

st.subheader("Downloads:")
generate_excel_download_link(df_selected)

