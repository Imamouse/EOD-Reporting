import streamlit as st #pip install streamlit
import pandas as pd #pip install pandas
import plotly.express as px #pip install plotly-express
import numpy as np #pip install
import base64 #Standard python module
from io import StringIO, BytesIO # Standard python module
import time
import json
import requests #pip install requests
from streamlit_lottie import st_lottie # pip install streamlit-lottie

st.set_page_config(page_title="Excel Plotter", page_icon=":card_file_box:", layout="wide")
st.title("Excel Plotter 🗃" )
st.subheader("Drag me an excel file.")

hide_menu_style = """
        <style>
        #MainMenu {visibility: hidden;}
	    footer {visibility: hidden;}
        </style>
        """
st.markdown(hide_menu_style, unsafe_allow_html=True)

def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "rts_rto" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link1(df1):
    towrite = BytesIO()
    df1.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Gcash" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link2(df2):
    towrite = BytesIO()
    df2.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Shopee" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link3(df3):
    towrite = BytesIO()
    df3.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Paymaya" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link4(df4):
    towrite = BytesIO()
    df4.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Reversal" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link5(df5):
    towrite = BytesIO()
    df5.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Monetary_appeasement" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link6(df6):
    towrite = BytesIO()
    df6.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "iPay88" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_excel_download_link7(df7):
    towrite = BytesIO()
    df7.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    new_name = "Project Grow" + "_" + time.strftime("%d/%m/%Y") + '.xls'
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{new_name}.xlsx"> Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

try:
    uploaded_file = st.file_uploader("Choose a xslx file", )
    uploaded_file = pd.ExcelFile(uploaded_file)
    df = pd.read_excel(uploaded_file, sheet_name="REFUND_RTS.RTO ", engine="openpyxl")
    df1 = pd.read_excel(uploaded_file, sheet_name="GCASH_THIRD PARTY", engine="openpyxl")
    df2 = pd.read_excel(uploaded_file, sheet_name="SHOPEE_THIRD PARTY", engine="openpyxl")
    df3 = pd.read_excel(uploaded_file, sheet_name="PAYMAYA_THIRD PARTY", engine="openpyxl")
    df4 = pd.read_excel(uploaded_file, sheet_name="REVERSAL", engine="openpyxl")
    df5 = pd.read_excel(uploaded_file, sheet_name="Monetary appeasement", engine="openpyxl")
    df6 = pd.read_excel(uploaded_file, sheet_name="iPay88", engine="openpyxl")
    df7 = pd.read_excel(uploaded_file, sheet_name="Project Grow (Re-Open Banner)", engine="openpyxl")

    st.sidebar.header("Please Filter Here:")
    status = st.sidebar.multiselect(
        "Select the RTS_RTO Status:",
        options=df["Status"].unique(),
        default=df["Status"].unique(),
    )

    status1 = st.sidebar.multiselect(
        "Select the GCash Status:",
        options=df1["Status"].unique(),
        default=df1["Status"].unique(),
    )

    status2 = st.sidebar.multiselect(
        "Select the Shopee Status:",
        options=df2["Status"].unique(),
        default=df2["Status"].unique(),
    )

    status3 = st.sidebar.multiselect(
        "Select the Paymaya Status:",
        options=df3["Status"].unique(),
        default=df3["Status"].unique(),
    )

    status4 = st.sidebar.multiselect(
        "Select the Reversal Status:",
        options=df4["Status"].unique(),
        default=df4["Status"].unique(),
    )

    status5 = st.sidebar.multiselect(
        "Select the Monetary Appeasement Status:",
        options=df5["Status"].unique(),
        default=df5["Status"].unique(),
    )

    status6 = st.sidebar.multiselect(
        "Select the iPay88 Status:",
        options=df6["Status"].unique(),
        default=df6["Status"].unique(),
    )

    status7 = st.sidebar.multiselect(
        "Select the Project Grow (Re-Open Banner) Status:",
        options=df7["Status"].unique(),
        default=df7["Status"].unique(),
    )

    df_selection = df.query (
        "Status == @status"
    )

    df_selection1 = df1.query (
        "Status == @status1"
    )

    df_selection2 = df2.query (
        "Status == @status2"
    )

    df_selection3 = df3.query (
        "Status == @status3"
    )

    df_selection4 = df4.query (
        "Status == @status4"
    )

    df_selection5 = df5.query (
        "Status == @status5"
    )

    df_selection6 = df6.query (
        "Status == @status6"
    )

    df_selection7 = df7.query (
        "Status == @status7"
    )



    df_selected = df_selection.drop(columns=["Entry Date",
                    "Date ticket was handled by Inspiro Esca Team (Date 1st Escalated)2",
                    "Email",
                    "Delivery Address",
                    "Date of order transaction",
                    "Remarks (To be filled by E-channel Team)",
                    "ESCA Comment", 
                    "Date refunded (To be filled by E-channel Team)"               
    ])

    df_selected1 = df_selection1.drop(columns=["Entry Date",
                    "Alternate Contact No.",
                    "EChannel Remarks",
                    "EC Status",
                    "PARTNER REMARKS", 
                    "ESCA Comments",
                    "Date Refunded by Partner"               
    ])

    df_selected2 = df_selection2.drop(columns=["Entry Date",
                    "Alternate Contact No.",
                    "EC Status",
                    "IT Comments",
                    "EChannel Remarks",
                    "PARTNER REMARKS", 
                    "Date Refunded by Partner",
                    "Escalations Remarks",           
    ])

    df_selected3 = df_selection3.drop(columns=["Date Logged",
                    "Alternate Contact Number",
                    "MT Remarks",
                    "IT Comments",
                    "MT Status",          
    ])

    df_selected4 = df_selection4.drop(columns=["Entry Date",
                    "IT Remarks (from the INT ticket)",
                    "Date of endorsement",
                    "Remarks (To be filled by E-channel Team)",
                    "Status (To be filled by E-channel Team)",
                    "Date refunded (To be filled by E-channel Team)"   

    ])

    df_selected5 = df_selection5.drop(columns=["Entry Date",
                    "Alternate Contact No.",
                    "EChannel Remarks",
                    "ESCA Comments",
                    "EC Status",
                    "PARTNER REMARKS",
                    "Date Refunded by Partner"
    ])

    df_selected6 = df_selection6.drop(columns=["Entry Date",
                    "IT Remarks",
                    "EChannel Remarks",
    ])

    df_selected7 = df_selection7.drop(columns=["Entry date",
                "Receive name",
                "Echannel Remarks"                          
    ])

    df_selected = df_selected.style.format({"Delivery no.": lambda x : '{0}'.format(x)})
    df_selected1 = df_selected1.style.format({"Dito  Number": lambda x : '{0}'.format(x)})
    df_selected2 = df_selected2.style.format({"Dito  Number": lambda x : '{0}'.format(x)})
    df_selected3 = df_selected3.style.format({"Dito  Number": lambda x : '{0}'.format(x)})
    df_selected4 = df_selected4.style.format({"Dito  Number": lambda x : '{0}'.format(x)})
    df_selected5 = df_selected5.style.format({"Dito  Number": lambda x : '{0}'.format(x)})
    df_selected6 = df_selected6.style.format({"Min": lambda x : '{0}'.format(x)})
    df_selected7 = df_selected7.style.format({"DITO Number": lambda x : '{0}'.format(x)})

    st.markdown("----")
    st.markdown("###")
    st.subheader(" RTS_RTO")
    st.dataframe(df_selected)
    st.subheader("Download RTS_RTO:")
    generate_excel_download_link(df_selected)

    st.markdown("----")
    st.markdown("###")
    st.subheader(" GCash")
    st.dataframe(df_selected1)
    st.subheader("Download GCash:")
    generate_excel_download_link1(df_selected1)

    st.markdown("----")
    st.subheader(" Shopee")
    st.dataframe(df_selected2)
    st.subheader("Download Shopee:")
    generate_excel_download_link2(df_selected2)

    st.markdown("----")
    st.subheader(" Paymaya")
    st.dataframe(df_selected3)
    st.subheader("Download Paymaya:")
    generate_excel_download_link3(df_selected3)

    st.markdown("----")
    st.subheader(" Reversal")
    st.dataframe(df_selected4)
    st.subheader("Download Reversal:")
    generate_excel_download_link4(df_selected4)

    st.markdown("----")
    st.subheader(" Monetary Appeasement")
    st.dataframe(df_selected5)
    st.subheader("Download Monetary Appeasement:")
    generate_excel_download_link5(df_selected5)

    st.markdown("----")
    st.subheader(" iPay88")
    st.dataframe(df_selected6)
    st.subheader("Download iPay88:")
    generate_excel_download_link6(df_selected6)

    st.markdown("----")
    st.subheader(" Project Grow (Re-Open Banner)")
    st.dataframe(df_selected7)
    st.subheader("Download Project Grow (Re-Open Banner):")
    generate_excel_download_link7(df_selected7)

except:
    def load_lottieurl(url: str):
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()


    excel_lottie = load_lottieurl("https://assets2.lottiefiles.com/packages/lf20_a4xidk9x.json")
    
    _left, mid, _right = st.columns(3)
    with mid:
        st_lottie(
            excel_lottie,
            loop=True,
            quality="low",
    )