import streamlit as st

pages = st.source_util.get_pages('Welcome.py')
new_page_names = {
  'Template': '📄 Template',
  'Visualization': '📊 Visualization',
}

for key, page in pages.items():
  if page['page_name'] in new_page_names:
    page['page_name'] = new_page_names[page['page_name']]
    
st.set_page_config(
    page_title="Cat Input File Application"
)

st.write("# Welcome to the Cat Input Application! 👋")


st.markdown('This application aims to a simple and efficient way to help you handling CAT data by giving you 2 options :')
st.markdown('- An **Excel Template** depends on your criteria (Type, Peril, Currency)')
st.markdown('- A **vizualisation** of your data with some graphs and tables')
st.markdown('**REMINDER : The Cat Input Application is still in Proof of Concept, any feedback will be more than welcomed to improve the App !** ')
    