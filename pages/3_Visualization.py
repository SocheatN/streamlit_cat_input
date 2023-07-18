import streamlit as st
#from st_aggrid import AgGrid
import pandas as pd
import numpy as np
import plotly.express as px
import base64
from io import StringIO, BytesIO

st.set_page_config(page_title="Visualization", page_icon=":egg:", layout="wide")

#st.markdown("# CAT Data Visualization")

# ---- DEF ----
@st.cache_data(show_spinner=False)
def convert_into_csv(df):
    return df.to_csv(index = False).encode('utf-8')

def generate_download_button(csv_data, file_label):
    st.download_button(label=f"Download {file_label} as CSV",
                           data=csv_data,
                           file_name=f"data_{file_label}.csv",
                           mime='text/csv')


@st.cache_data(show_spinner=False)
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
@st.cache_data(show_spinner=False)
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
        
local_css("style.css")

# ---- HEADER SECTION ----
with st.container():    
    st.title("Cat Input File Visualization")
    st.subheader("Upload your CAT Data")
    
# ---- SELECTION ----


# First selection : RMS or AIR ? 
selection1 = st.radio('Type : ', ['RMS', 'AIR'])

# Second selection : EQ or TC ? 
selection2 = st.write('Peril : ')
selection2_option_1 = st.checkbox('EQ')
selection2_option_2 = st.checkbox('TC')

# Third Selection : Currency ?

list_currency = sorted(['AUD', 'CNY', 'IDR', 'INR', 'JPY', 'NZD', 'PHP', 'TWD', 'USD', 'EUR', 'KRW', 'HKD'])
list_currency.append('Other Currency')
currency = st.selectbox("Select currency", options=list_currency)

if currency == "Other Currency": 
    currency_otherOption = st.text_input("Other Currency: ")
    
uploaded_file = st.file_uploader('Choose an excel file')
if uploaded_file:
    st.markdown('---')
    
    df_Account = pd.read_excel(uploaded_file,sheet_name='Account Group')
    df_EQ = pd.read_excel(uploaded_file,sheet_name='Exp_EQ')
    df_TC = pd.read_excel(uploaded_file,sheet_name='Exp_TC')
    df_Occ = pd.read_excel(uploaded_file,sheet_name='Occ')
    df_Cons = pd.read_excel(uploaded_file,sheet_name='Cons')
    df_BH = pd.read_excel(uploaded_file,sheet_name='BH')
    df_YB = pd.read_excel(uploaded_file,sheet_name='YB')
               
        
    # -- GROUP DATAFRAME
    
    if selection1 == 'RMS':
        groupby_column = st.selectbox(
             'What would you like to group by?',
            ('ACCNTNUM', 'STATE', 'DISTRICT'),
        )
        output_columns = ['TIV', 'SITELIM']
        
        
        
        
        df_grouped1 = df_EQ.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
        df_grouped2 = df_TC.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
        
        # -- PLOT DATAFRAME
    
        
        fig1 = px.bar(
            df_grouped1,
            x=groupby_column,
            y='TIV',
            color='SITELIM',
            color_continuous_scale=['green', 'yellow', 'red'],
            template='plotly_white',
            title=f'<b>EQ:TIV & Sitelimit by {groupby_column}</b>',
            text_auto = True
        )
        fig1.update_layout(xaxis={'categoryorder':'total descending'})
        
        fig2 = px.bar(
            df_grouped2,
            x=groupby_column,
            y='TIV',
            color='SITELIM',
            color_continuous_scale=['green', 'yellow', 'red'],
            template='plotly_white',
            title=f'<b>TC:TIV & Sitelimit by {groupby_column}</b>',
            text_auto = True
        )
        fig2.update_layout(xaxis={'categoryorder':'total descending'})
        
        
        fig3 = px.bar(
            df_Occ,
            x='LOBNAME',
            y='Occ_split',
            color="Occupancy",
            color_continuous_scale='BuGn',
            template='plotly_white',
            title=f'<b>Split by Occupancy</b>',
            text = df_Occ['Occ_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )   
        
        
        fig4 = px.bar(
            df_Cons,
            x='LOBNAME',
            y='Cons_split',
            color="Construction",
            color_continuous_scale='Blues',
            template='plotly_white',
            title=f'<b>Split by Construction</b>',
            text = df_Cons['Cons_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
        
        fig5 = px.bar(
            df_BH,
            x='LOBNAME',
            y='BH_split',
            color="BH",
            color_continuous_scale='Purples',
            template='plotly_white',
            title=f'<b>Split by Building Height</b>',
            text = df_BH['BH_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
            
        fig6 = px.bar(
            df_YB,
            x='LOBNAME',
            y='YB_split',
            color="YB",
            color_continuous_scale='Oranges',
            template='plotly_white',
            title=f'<b>Split by Year Built</b>',
            text = df_YB['YB_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
           
        
        if currency == 'Other Currency':
            cur = currency_otherOption
        else:
            cur = currency    
        
        EQ_loc_file=loc_file_RMS(df_EQ,'EQ',df_Occ,df_Cons,df_BH, cur)
        TC_loc_file=loc_file_RMS(df_TC,'WS',df_Occ,df_Cons,df_BH, cur)
        
        if selection2_option_1 and selection2_option_2:  # EQ and TC selected

            # -- PLOT THE GRAPHS
            Plot1, Plot2 = st.columns(2)
            with Plot1:
                st.plotly_chart(fig1)
                
            with Plot2:
                st.plotly_chart(fig2)
                
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)

            # -- TABLES FOR ACCOUNT/LOCATION            
            df_Location=pd.concat([EQ_loc_file,TC_loc_file])
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("RMS: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("RMS: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            csv_Account = convert_into_csv(edited_df_Account)
            csv_Location = convert_into_csv(edited_df_Location)
            
            # -- DOWNLOAD SECTION
            st.subheader('Downloads:')
            generate_download_button(csv_Account, 'Account')
            generate_download_button(csv_Location, 'Location')
            
                    
        if selection2_option_1 and not selection2_option_2:  # Only EQ selected

            # -- PLOT THE GRAPHS

            st.plotly_chart(fig1)           
                
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)


            # -- TABLES FOR ACCOUNT/LOCATION                        
            df_Location=EQ_loc_file
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("RMS: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("RMS: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            csv_Account = convert_into_csv(edited_df_Account)
            csv_Location = convert_into_csv(edited_df_Location)
            
            # -- DOWNLOAD SECTION
            st.subheader('Downloads:')
            generate_download_button(csv_Account, 'Account')
            generate_download_button(csv_Location, 'Location')
            
        if selection2_option_2 and not selection2_option_1:  # Only TC selected

            # -- PLOT THE GRAPHS
            st.plotly_chart(fig2)
                           
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)

            # -- TABLES FOR ACCOUNT/LOCATION            
            df_Location=TC_loc_file
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("RMS: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("RMS: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            csv_Account = convert_into_csv(edited_df_Account)
            csv_Location = convert_into_csv(edited_df_Location)
            
            # -- DOWNLOAD SECTION
            st.subheader('Downloads:')
            generate_download_button(csv_Account, 'Account')
            generate_download_button(csv_Location, 'Location')
                   
        
    if selection1 == 'AIR':
        hide_dataframe_row_index = """
                    <style>
                    .row_heading.level0 {display:none}
                    .blank {display:none}
                    </style>
                    """
        
        # Inject CSS with Markdown
        st.markdown(hide_dataframe_row_index, unsafe_allow_html=True)
        
        groupby_column = st.selectbox(
             'What would you like to group by?',
            ('UDF1', 'Area'),
        )
        output_columns = ['BuildingValue', 'ContentsValue','TimeElementValue', 'OtherValue']
        
        df_grouped1 = df_EQ.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
        df_grouped2 = df_TC.groupby(by=[groupby_column], as_index=False)[output_columns].sum()
        
        # -- PLOT DATAFRAME
    
        
        fig1 = px.bar(
            df_grouped1,
            x=groupby_column,
            y='BuildingValue',
            color='ContentsValue',
            color_continuous_scale=['green', 'yellow', 'red'],
            template='plotly_white',
            title=f'<b>EQ:BuildingValue & ContentsValue by {groupby_column}</b>',
            text_auto = True
        )
        fig1.update_layout(xaxis={'categoryorder':'total descending'})
        
        fig2 = px.bar(
            df_grouped2,
            x=groupby_column,
            y='BuildingValue',
            color='ContentsValue',
            color_continuous_scale=['green', 'yellow', 'red'],
            template='plotly_white',
            title=f'<b>TC:BuildingValue & ContentsValue by {groupby_column}</b>',
            text_auto = True
        )
        fig2.update_layout(xaxis={'categoryorder':'total descending'})
        
        
        fig3 = px.bar(
            df_Occ,
            x='UDF1',
            y='Occ_split',
            color="Occupancy",
            color_continuous_scale='BuGn',
            template='plotly_white',
            title=f'<b>Split by Occupancy</b>',
            text = df_Occ['Occ_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )   
        
        
        fig4 = px.bar(
            df_Cons,
            x='UDF1',
            y='Cons_split',
            color="ConstructionCategory",
            color_continuous_scale='Blues',
            template='plotly_white',
            title=f'<b>Split by Construction</b>',
            text = df_Cons['Cons_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
        
        fig5 = px.bar(
            df_BH,
            x='UDF1',
            y='BH_split',
            color="NumberOfStories",
            color_continuous_scale='Purples',
            template='plotly_white',
            title=f'<b>Split by Building Height</b>',
            text = df_BH['BH_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
            
        fig6 = px.bar(
            df_YB,
            x='UDF1',
            y='YB_split',
            color="YearBuilt",
            color_continuous_scale='Oranges',
            template='plotly_white',
            title=f'<b>Split by Year Built</b>',
            text = df_YB['YB_split'].apply(lambda x: '{0:1.2f}%'.format(x*100))
        )
        
        
        if currency == 'Other Currency':
            cur = currency_otherOption
        else:
            cur = currency
            
        EQ_loc_file=loc_file_AIR(df_EQ,'EQ',df_Occ,df_Cons,df_BH, cur)
        TC_loc_file=loc_file_AIR(df_TC,'WS',df_Occ,df_Cons,df_BH, cur)
                
            
        if selection2_option_1 and selection2_option_2:  # EQ and TC selected

            # -- PLOT THE GRAPHS 
            Table1, Table2 = st.columns(2)
            with Table1:                
                st.header('EQ table')
                st.table(df_grouped1.style.hide(axis="index"))
                
            with Table2:
                st.header('TC Table')
                st.table(df_grouped2.style.hide(axis="index"))
                        
            
            Plot1, Plot2 = st.columns(2)
            with Plot1:                
                st.plotly_chart(fig1)
                
            with Plot2:
                st.plotly_chart(fig2)
                
                
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)
                     
            # -- TABLES FOR ACCOUNT/LOCATION
            
            df_Location=pd.concat([EQ_loc_file,TC_loc_file])
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("AIR: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("AIR: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            # -- DOWNLOAD SECTION
            st.write('Download the tables ?')
            confirm_button = st.checkbox('Yes')
            if confirm_button:                
                st.subheader('Downloads:')
                with st.spinner('Wait for it...'):
                    generate_excel_download_link(edited_df_Account,'AIR input file: Account')
                    generate_excel_download_link(edited_df_Location,'AIR input file: Location')
                    #generate_html_download_link(fig1)
                    #generate_html_download_link(fig2)
                    st.success('Done!')
                    
        if selection2_option_1 and not selection2_option_2:  # Only EQ selected

            # -- PLOT THE GRAPHS

            st.plotly_chart(fig1)
                
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)

            # -- TABLES FOR ACCOUNT/LOCATION
                        
            df_Location=EQ_loc_file
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("AIR: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("AIR: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            # -- DOWNLOAD SECTION
            st.write('Download the tables ?')
            confirm_button = st.checkbox('Yes')
            if confirm_button:                
                st.subheader('Downloads:')
                with st.spinner('Wait for it...'):
                    generate_excel_download_link(edited_df_Account,'AIR input file: Account')
                    generate_excel_download_link(edited_df_Location,'AIR input file: Location')
                    #generate_html_download_link(fig1)
                    #generate_html_download_link(fig2)
                    st.success('Done!')
                    
        if selection2_option_2 and not selection2_option_1:  # Only TC selected

            # -- PLOT THE GRAPHS

            st.plotly_chart(fig2)
                
            Plot3, Plot4, Plot5, Plot6 = st.columns(4)
            with Plot3:
                st.plotly_chart(fig3)
            with Plot4:
                st.plotly_chart(fig4)
            with Plot5:
                st.plotly_chart(fig5)
            with Plot6:
                st.plotly_chart(fig6)
                
            # -- TABLES FOR ACCOUNT/LOCATION    
            
            df_Location=TC_loc_file
            df_Location['LOCNUM']=df_Location.reset_index().index+1
                            
            Account, Location = st.columns(2)
            with Account:
                st.header("AIR: Account file")
                edited_df_Account = st.experimental_data_editor(df_Account, num_rows = 'dynamic')
            with Location:
                st.header("AIR: Location file")
                edited_df_Location = st.experimental_data_editor(df_Location, num_rows = 'dynamic')
            
            # -- DOWNLOAD SECTION
            st.write('Download the tables ?')
            confirm_button = st.checkbox('Yes')
            if confirm_button:                
                st.subheader('Downloads:')
                with st.spinner('Wait for it...'):
                    generate_excel_download_link(edited_df_Account,'AIR input file: Account')
                    generate_excel_download_link(edited_df_Location,'AIR input file: Location')
                    #generate_html_download_link(fig1)
                    #generate_html_download_link(fig2)
                    st.success('Done!')
