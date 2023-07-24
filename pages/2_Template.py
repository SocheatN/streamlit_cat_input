import streamlit as st
#from st_aggrid import AgGrid
import pandas as pd
import numpy as np
import plotly.express as px
import base64
from io import StringIO, BytesIO
from PIL import Image
#import time
#from openpyxl import load_workbook

st.set_page_config(page_title="Template", page_icon=":egg:", layout="wide")

#st.markdown("# CAT Excel Template")

# ---- DEF ----   
@st.cache_data(show_spinner=False, experimental_allow_widgets=True)
def generate_template_RMS(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'RMS_Account Group')
            df2 = pd.read_excel(xls, 'RMS_Exp_EQ')
            df3 = pd.read_excel(xls, 'RMS_Exp_TC')
            df4 = pd.read_excel(xls, 'RMS_Occ')
            df5 = pd.read_excel(xls, 'RMS_Cons')
            df6 = pd.read_excel(xls, 'RMS_BH')
            df7 = pd.read_excel(xls, 'RMS_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_EQ', index = False)
                df3.to_excel(writer, sheet_name='EXP_TC', index = False)
                df4.to_excel(writer, sheet_name='Occ', index = False)
                df5.to_excel(writer, sheet_name='Cons', index = False)
                df6.to_excel(writer, sheet_name='BH', index = False)
                df7.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (RMS)",
                    data=buffer,
                    file_name="RMS_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)                
def generate_template_RMS_EQ(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'RMS_Account Group')
            df2 = pd.read_excel(xls, 'RMS_Exp_EQ')
            df3 = pd.read_excel(xls, 'RMS_Occ')
            df4 = pd.read_excel(xls, 'RMS_Cons')
            df5 = pd.read_excel(xls, 'RMS_BH')
            df6 = pd.read_excel(xls, 'RMS_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_EQ', index = False)
                df3.to_excel(writer, sheet_name='Occ', index = False)
                df4.to_excel(writer, sheet_name='Cons', index = False)
                df5.to_excel(writer, sheet_name='BH', index = False)
                df6.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (RMS - Only EQ)",
                    data=buffer,
                    file_name="RMS_EQ_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)                
def generate_template_RMS_TC(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'RMS_Account Group')
            df2 = pd.read_excel(xls, 'RMS_Exp_TC')
            df3 = pd.read_excel(xls, 'RMS_Occ')
            df4 = pd.read_excel(xls, 'RMS_Cons')
            df5 = pd.read_excel(xls, 'RMS_BH')
            df6 = pd.read_excel(xls, 'RMS_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_TC', index = False)
                df3.to_excel(writer, sheet_name='Occ', index = False)
                df4.to_excel(writer, sheet_name='Cons', index = False)
                df5.to_excel(writer, sheet_name='BH', index = False)
                df6.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (RMS - Only TC)",
                    data=buffer,
                    file_name="RMS_TC_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)                
def generate_template_AIR(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'AIR_Account Group')
            df2 = pd.read_excel(xls, 'AIR_Exp_EQ')
            df3 = pd.read_excel(xls, 'AIR_Exp_TC')
            df4 = pd.read_excel(xls, 'AIR_Occ')
            df5 = pd.read_excel(xls, 'AIR_Cons')
            df6 = pd.read_excel(xls, 'AIR_BH')
            df7 = pd.read_excel(xls, 'AIR_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_EQ', index = False)
                df3.to_excel(writer, sheet_name='EXP_TC', index = False)
                df4.to_excel(writer, sheet_name='Occ', index = False)
                df5.to_excel(writer, sheet_name='Cons', index = False)
                df6.to_excel(writer, sheet_name='BH', index = False)
                df7.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (AIR)",
                    data=buffer,
                    file_name="AIR_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)                
def generate_template_AIR_EQ(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'AIR_Account Group')
            df2 = pd.read_excel(xls, 'AIR_Exp_EQ')
            df3 = pd.read_excel(xls, 'AIR_Occ')
            df4 = pd.read_excel(xls, 'AIR_Cons')
            df5 = pd.read_excel(xls, 'AIR_BH')
            df6 = pd.read_excel(xls, 'AIR_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_EQ', index = False)
                df3.to_excel(writer, sheet_name='Occ', index = False)
                df4.to_excel(writer, sheet_name='Cons', index = False)
                df5.to_excel(writer, sheet_name='BH', index = False)
                df6.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (AIR - Only EQ)",
                    data=buffer,
                    file_name="AIR_EQ_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)                
def generate_template_AIR_TC(excel_template):
            buffer = BytesIO()
            xls = pd.ExcelFile(excel_template)
            df1 = pd.read_excel(xls, 'AIR_Account Group')
            df2 = pd.read_excel(xls, 'AIR_Exp_TC')
            df3 = pd.read_excel(xls, 'AIR_Occ')
            df4 = pd.read_excel(xls, 'AIR_Cons')
            df5 = pd.read_excel(xls, 'AIR_BH')
            df6 = pd.read_excel(xls, 'AIR_YB')
            
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Write each dataframe to a different worksheet.
                df1.to_excel(writer, sheet_name='Account Group', index = False)
                df2.to_excel(writer, sheet_name='EXP_TC', index = False)
                df3.to_excel(writer, sheet_name='Occ', index = False)
                df4.to_excel(writer, sheet_name='Cons', index = False)
                df5.to_excel(writer, sheet_name='BH', index = False)
                df6.to_excel(writer, sheet_name='YB', index = False)
                
                # Close the Pandas Excel writer and output the Excel file to the buffer
                writer.save()
            
                st.download_button(
                    label="Download Excel Template (AIR - Only TC)",
                    data=buffer,
                    file_name="AIR_TC_Template.xlsx")

@st.cache_data(show_spinner=False, experimental_allow_widgets=True)
def tutorial_download():
    st.markdown('To select the destination folder for your template, make sure to have the right settings in your browser.')
    image1, image2, image3 = st.columns(3)
    with image1:
        image1 = Image.open('MicrosoftEdge_Download.png')
        st.image(image1, caption='Microsft Edge')
    with image2:
        image2 = Image.open('Chrome_Download.png')
        st.image(image2, caption='Google Chrome')
    with image3:
        image3 = Image.open('Firefox_Download.png')
        st.image(image3, caption='Mozilla Firefox')

# Use local CSS
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
        


local_css("style.css")

# ---- HEADER SECTION ----
with st.container():    
    st.title("Cat Input File Generator")    
    st.markdown('IMPORTANT : Make sure to fill correctly the feature "LOBNAME" in your Template')
    st.subheader("Generate your Excel Template")

# ---- SELECTION ----


# First selection : RMS or AIR ? 
selection1 = st.radio('Type : ', ['RMS', 'AIR'])

# Second selection : EQ or TC ? 
selection2 = st.write('Peril : ')
selection2_option_1 = st.checkbox('EQ')
selection2_option_2 = st.checkbox('TC')

# --- TEMPLATE SECTION -----
if selection1 == 'RMS':
    if selection2_option_1 and selection2_option_2:   # EQ and TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        
        with st.spinner('Wait for it...'):
                generate_template_RMS('Template.xlsx')
                st.success('Done!')   
        tutorial_download()
            
    if selection2_option_1 and not selection2_option_2:  # Only EQ selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_RMS_EQ('Template.xlsx')
                st.success('Done!')
        tutorial_download()
                
    if selection2_option_2 and not selection2_option_1:  # Only TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_RMS_TC('Template.xlsx')
                st.success('Done!')
        tutorial_download()

if selection1 == 'AIR':
    if selection2_option_1 and selection2_option_2:  # EQ and TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR('Template.xlsx')
                st.success('Done!')
        tutorial_download()
            
    if selection2_option_1 and not selection2_option_2: # Only EQ selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR_EQ('Template.xlsx')
                st.success('Done!')
        tutorial_download()
                
    if selection2_option_2 and not selection2_option_1: # Only TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR_TC('Template.xlsx')
                st.success('Done!')  
        tutorial_download()