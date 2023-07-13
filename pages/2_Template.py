import streamlit as st
#from st_aggrid import AgGrid
import pandas as pd
import numpy as np
import plotly.express as px
import base64
from io import StringIO, BytesIO
#import time
#from openpyxl import load_workbook

st.set_page_config(page_title="Template", page_icon=":egg:", layout="wide")

#st.markdown("# CAT Excel Template")

# ---- DEF ----
def generate_excel_download_link(df, file_name):
    towrite = BytesIO()
    df.to_excel(towrite,  index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download '+file_name+' </a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)

def loc_file_RMS(Exp,peril,df_Occ,df_Cons,df_BH, currency):
    df=Exp.merge(df_Occ,on='LOBNAME').merge(df_Cons,on='LOBNAME').merge(df_BH,on='LOBNAME').merge(df_YB,on='LOBNAME') 
    df['EQCV1VAL'],df['EQCV2VAL'],df['EQCV3VAL'],df['EQSITELIM'], df['WSCV1VAL'],df['WSCV2VAL'],df['WSCV3VAL'],df['WSSITELIM']=0,0,0,0,0,0,0,0
    df[peril+'CV1VAL']=df['BLDG']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df[peril+'CV2VAL']=df['CONT']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df[peril+'CV3VAL']=df['BI']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df['NUMBLDGS']=np.ceil(df['NUMBLDGS']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split'])
    df[peril+'SITELIM']=df['SITELIM']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df=df[df[peril+'CV1VAL']+df[peril+'CV2VAL']+df[peril+'CV3VAL']>0]
    
    df['ACCNTNUM'],df['CNTRYCODE'],df['CNTRYSCHEME'],df['EQCV1VCUR'],df['EQCV2VCUR'],df['EQCV3VCUR'], df['EQSITELCUR'],df['WSCV1VCUR'],df['WSCV2VCUR'],df['WSCV3VCUR'],df['WSSITELCUR']= df['LOBNAME']+'_'+peril,'CN','ISO2A', currency, currency, currency, currency, currency, currency, currency, currency
    
    loc_file=df.drop(columns = ['BLDG','CONT','BI','TIV','SITELIM','Occupancy','Occ_split','Construction','Cons_split','BH','BH_split','YB','YB_split'])
    return loc_file

def loc_file_AIR(Exp,peril,df_Occ,df_Cons,df_BH, currency):
    df=Exp.merge(df_Occ,on='ContractID').merge(df_Cons,on='ContractID').merge(df_BH,on='ContractID').merge(df_YB,on='ContractID') 
    df['EQCV1VAL'],df['EQCV2VAL'],df['EQSITELIM'], df['WSCV1VAL'],df['WSCV2VAL'],df['WSSITELIM']=0,0,0,0,0,0
    df[peril+'CV1VAL']=df['BuildingValue']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df[peril+'CV2VAL']=df['ContentsValue']*df['Occ_split']*df['Cons_split']*df['BH_split']*df['YB_split']
    df=df[df[peril+'CV1VAL']+df[peril+'CV2VAL']>0]
    
    df['ContractID'], df['CNTRYCODE'], df['CNTRYSCHEME'], df['EQCV1VCUR'], df['EQCV2VCUR'], df['EQSITELCUR'], df['WSCV1VCUR'], df['WSCV2VCUR'], df['WSSITELCUR']= df['ContractID'], 'CN', 'ISO2A', currency, currency, currency, currency, currency, currency
    df = df.loc[:,~df.columns.duplicated()].copy()
    loc_file=df.drop(columns = ['BuildingValue','ContentsValue','Occupancy','Occ_split','Cons_split','BH_split','YB_split'])
    return loc_file

# Use local CSS
def local_css(file_name):
    with open(file_name) as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
        
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



local_css("style.css")

# ---- HEADER SECTION ----
with st.container():    
    st.title("Cat Input File Generator")
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
                generate_template_RMS('v7Template.xlsx')
                st.success('Done!')
            
    if selection2_option_1 and not selection2_option_2:  # Only EQ selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_RMS_EQ('v7Template.xlsx')
                st.success('Done!')
                
    if selection2_option_2 and not selection2_option_1:  # Only TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_RMS_TC('v7Template.xlsx')
                st.success('Done!')

if selection1 == 'AIR':
    if selection2_option_1 and selection2_option_2:  # EQ and TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR('v7Template.xlsx')
                st.success('Done!')
            
    if selection2_option_1 and not selection2_option_2: # Only EQ selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR_EQ('v7Template.xlsx')
                st.success('Done!')
                
    if selection2_option_2 and not selection2_option_1: # Only TC selected

        # -- DOWNLOAD SECTION
        st.subheader('Downloads:')
        with st.spinner('Wait for it...'):
                generate_template_AIR_TC('v7Template.xlsx')
                st.success('Done!')                