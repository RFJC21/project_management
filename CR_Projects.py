import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from st_aggrid import AgGrid
from st_aggrid import GridOptionsBuilder, AgGrid, GridUpdateMode, DataReturnMode
import xlsxwriter
import plotly.figure_factory as ff


st.set_page_config(page_title='CR Projects', layout='wide')

#st.session_state

# Load the data into a dataframe
FILEPATH = "files/projects_2023.xlsx"
df = pd.read_excel(FILEPATH, engine='openpyxl', sheet_name='projects_2023')
df_fte = pd.read_excel(FILEPATH, engine='openpyxl', sheet_name='FTE')

# sidebar
project = st.sidebar.multiselect('Select Project to see detail:', df['Project'].unique())

# Create a Gantt chart using plotly
if project:
    df = df[df['Project'].isin(project)]
    fig = px.timeline(df,
                      x_start="Start_Date",
                      x_end="End_Date",
                      y="Task",
                      color="Project",
                      animation_group='Task',
                      text='Status',
                      title='CR Projects Timeline',
                      labels='Status',
                      color_continuous_scale=px.colors.sequential.Plasma,
                      #parent='Parent'
                      # animation_group='Owner'
                      )

    fig.update_yaxes(autorange="reversed")

    # Display the chart
    st.write('##')
    st.plotly_chart(fig)

    st.dataframe(df)
else:
    st.subheader('Timeline')
    fig = px.timeline(df,
                      x_start="Start_Date",
                      x_end="End_Date",
                      y="Project",
                      color="Owner",
                      animation_group='Task',
                      text='Resources',
                      #title='CR Projects Timeline',
                      color_continuous_scale=px.colors.sequential.Plasma)

    fig.update_yaxes(autorange="reversed")

    # Display the chart
    st.write('##')
    st.plotly_chart(fig)

    # edit the excel file with new inputs
    st.subheader('Edit tasks/subtasks')

    gd = GridOptionsBuilder.from_dataframe(df)
    gd.configure_pagination(enabled=True)
    gd.configure_default_column(editable=True, groupable=True)
    gd.configure_selection(selection_mode="multiple", use_checkbox=True)
    gridoptions = gd.build()
    grid_table = AgGrid(
        df,
        gridOptions=gridoptions,
        update_mode=GridUpdateMode.SELECTION_CHANGED,
        theme="streamlit")

    # grid_return = AgGrid(df, gridoptions)
    df_updated = grid_table["data"]

    if st.button('Submit Changes'):
        # update projects file
        with pd.ExcelWriter("files/projects_2023.xlsx") as writer:
            # use to_excel function and specify the sheet_name and index
            # to store the dataframe in specified sheet
            df_updated.to_excel(writer, sheet_name="projects_2023", index=False)
            df_fte.to_excel(writer, sheet_name="FTE", index=False)

    st.write('##')
    st.subheader('Add tasks/subtasks')
    if st.checkbox('Add new task/subtask'):
        with st.form('Form'):
            start_date = st.date_input('Start Date')
            end_date = st.date_input('End Date')
            project_input = st.text_input('Project')
            subtask = st.text_input('Subtask')
            owner = st.text_input('Owner')
            resources = st.number_input('Resources', 1,3,1)
            status = st.text_input('Status')
            comments = st.text_input('Notes')

            new_task = {'Start_Date': start_date,
                        'End_Date': end_date,
                        'Project': project_input,
                        'Task': subtask,
                        'Owner': owner,
                        'Resources': resources,
                        'Status': status,
                        'Comments': comments}

            submit = st.form_submit_button('Add task/subtask')

            if submit:
                df = df.append(new_task, ignore_index=True)

                # update projects file
                with pd.ExcelWriter("files/projects_2023.xlsx") as writer:
                    # use to_excel function and specify the sheet_name and index
                    # to store the dataframe in specified sheet
                    df.to_excel(writer, sheet_name="projects_2023", index=False)
                    df_fte.to_excel(writer, sheet_name="FTE", index=False)

# download table updated from server
st.sidebar.write('##')
st.sidebar.download_button(label='📥 Download Tasks', data=df.to_csv(), mime='text/csv')

