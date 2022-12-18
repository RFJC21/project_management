import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode
import xlsxwriter
import plotly.figure_factory as ff

st.set_page_config(page_title='CR Projects', layout='wide')
st.subheader('CR FTE')

#st.session_state

# Load the data into a dataframe
FILEPATH = "files/projects_2023.xlsx"
df = pd.read_excel(FILEPATH, engine='openpyxl', sheet_name='projects_2023')
df_fte = pd.read_excel(FILEPATH, engine='openpyxl', sheet_name='FTE')

df_fte_resume = df_fte.drop(df_fte.columns[[2,4,6,8,10,12]],axis = 1)

# edit the excel file with new inputs
gd = GridOptionsBuilder.from_dataframe(df_fte_resume)
gd.configure_pagination(enabled=True)
gd.configure_default_column(editable=True, groupable=True)
gd.configure_selection(selection_mode="multiple", use_checkbox=True)
gridoptions = gd.build()
grid_table_fte = AgGrid(
    df_fte_resume,
    gridOptions=gridoptions,
    update_mode=GridUpdateMode.SELECTION_CHANGED,
    theme="streamlit")

#grid_return = AgGrid(df, gridoptions)
df_updated = grid_table_fte["data"]

comment='''if st.button('Submit Changes'):
    # update projects file
    with pd.ExcelWriter("files/projects_2023.xlsx") as writer:
        # use to_excel function and specify the sheet_name and index
        # to store the dataframe in specified sheet
        df_updated.to_excel(writer, sheet_name="projects_2023", index=False)
        df_fte.to_excel(writer, sheet_name="FTE", index=False)'''

