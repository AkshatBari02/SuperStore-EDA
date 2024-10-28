import streamlit as st
import plotly.express as px
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")


st.set_page_config(page_title="SuperStore!!!", page_icon=":bar_chart:",layout="wide")

st.title(" :bar_chart: Sample Superstore EDA")
st.markdown("<style>.st-emotion-cache-12fmjuu{height:1rem;}</style>",unsafe_allow_html=True)
st.markdown('<style>div.block-container{padding-top:2rem;}</style>',unsafe_allow_html=True)



required_columns = ["Order Date", "Region", "State", "City", "Category", "Sales"]
fl = st.file_uploader(":file_folder: Upload a file",type=(['csv','txt','xlsx','xls']))
if fl is not None:
    filename = fl.name
    st.write(filename)
    if filename.endswith(('.csv', '.txt')):
        df = pd.read_csv(fl, encoding="ISO-8859-1")
    elif filename.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(fl)
else:
    os.chdir(r"G:\Data Analytics")
    df = pd.read_excel("Superstore.xls")

if not all(col in df.columns for col in required_columns):
    st.error("The uploaded file does not contain the required columns.")
else:
    col1, col2 = st.columns((2))
    df["Order Date"] = pd.to_datetime(df["Order Date"])

    # Getting min and max date
    startDate = pd.to_datetime(df["Order Date"]).min()
    endDate = pd.to_datetime(df["Order Date"]).max()

    with col1:
        date1 = pd.to_datetime(st.date_input("Start Date",startDate))
    with col2:
        date2 = pd.to_datetime(st.date_input("End Date",endDate))

    df = df[(df["Order Date"] >= date1) & (df["Order Date"] <= date2)].copy()

    st.sidebar.header("Choose your filter: ")

    # Create for Region
    region = st.sidebar.multiselect("Pick the region",df["Region"].unique())
    if not region:
        df2 = df.copy()
    else:
        df2 = df[df["Region"].isin(region)]

    # Create for State
    state = st.sidebar.multiselect("Pick the state",df["State"].unique())
    if not state:
        df3 = df2.copy()
    else:
        df3 = df2[df2["State"].isin(state)]


    # Create for City
    city = st.sidebar.multiselect("Pick the city",df3["City"].unique())


    # Filter the data based on Region,State,City

    if not region and not state and not city:
        filtered_df = df
    elif not state and not city:
        filtered_df = df[df["Region"].isin(region)]
    elif not region and not city:
        filtered_df = df[df["State"].isin(state)]
    elif state and city:
        filtered_df = df3[df["State"].isin(state) & df3["City"].isin(city)]
    elif region and city:
        filtered_df = df3[df["Region"].isin(region) & df3["City"].isin(city)]
    elif region and state:
        filtered_df = df3[df["Region"].isin(region) & df3["State"].isin(state)]
    elif city:
        filtered_df = df3[df3["City"].isin(city)]
    else:
        filtered_df = df3[df["Region"].isin(region) & df3["State"].isin(state) & df3["City"].isin(city)]

    category_df = filtered_df.groupby(by = ["Category"], as_index = False)["Sales"].sum()

    with col1:
        st.subheader("Category wise Sales")
        fig = px.bar(category_df, x = "Category", y="Sales", text = ['${:,.2f}'.format(x) for x in category_df["Sales"]], template="seaborn")
        st.plotly_chart(fig,use_container_width=True,height = 200)

    with col2:
        st.subheader("Region wise Sales")
        fig = px.pie(filtered_df,values = "Sales",names = "Region",hole = 0.5)
        fig.update_traces(text = filtered_df["Region"], textposition = "outside")
        st.plotly_chart(fig,use_container_width=True)


    cl1, cl2 = st.columns((2))

    with cl1:
        with st.expander("Category_ViewData"):
            st.write(category_df)
            
            # Download options
            download_type = st.selectbox("Select file format for download:", ['csv', 'xlsx'],key="category_format")
            
            if download_type == 'csv':
                csv = category_df.to_csv(index=False).encode('utf-8')
                st.download_button("Download as CSV file", data=csv, file_name="Category.csv", mime="text/csv", help='Click here to download data as CSV file')
            
            elif download_type == 'xlsx':
                excel_file = pd.ExcelWriter('Category.xlsx', engine='openpyxl')
                category_df.to_excel(excel_file, index=False, sheet_name='Category Data')
                excel_file.save()
                # Read the file to get bytes
                with open('Category.xlsx', 'rb') as f:
                    excel_data = f.read()
                st.download_button("Download as Excel file", data=excel_data, file_name="Category.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", help='Click here to download data as Excel file')

    with cl2:
        with st.expander("Region_ViewData"):
            region = filtered_df.groupby(by = "Region", as_index=False)["Sales"].sum()
            st.write(region)
            
            # Download options
            download_type = st.selectbox("Select file format for download:", ['csv', 'xlsx'],key="region_format")
            
            if download_type == 'csv':
                csv = category_df.to_csv(index=False).encode('utf-8')
                st.download_button("Download as CSV file", data=csv, file_name="Region.csv", mime="text/csv", help='Click here to download data as CSV file')
            
            elif download_type == 'xlsx':
                excel_file = pd.ExcelWriter('Region.xlsx', engine='openpyxl')
                category_df.to_excel(excel_file, index=False, sheet_name='Region Data')
                excel_file._save()
                # Read the file to get bytes
                with open('Region.xlsx', 'rb') as f:
                    excel_data = f.read()
                st.download_button("Download as Excel file", data=excel_data, file_name="Region.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", help='Click here to download data as Excel file')


    filtered_df["month_year"] = filtered_df["Order Date"].dt.to_period("M")
    st.subheader("Time Series Analysis")

    linechart = pd.DataFrame(filtered_df.groupby(filtered_df["month_year"].dt.strftime("%Y : %b"))["Sales"].sum()).reset_index()
    fig2 = px.line(linechart, x="month_year", y="Sales",labels={"Sales":"Amount"},height=500,width=1000,template="gridon")
    st.plotly_chart(fig2,use_container_width=True)

    with st.expander("View TimeSeries Data:"):
        st.write(linechart.T)
        # Download options
        download_type = st.selectbox("Select file format for download:", ['csv', 'xlsx'],key="timeSeries_format")
        
        if download_type == 'csv':
            csv = category_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download as CSV file", data=csv, file_name="TimeSeries.csv", mime="text/csv", help='Click here to download data as CSV file')
        
        elif download_type == 'xlsx':
            excel_file = pd.ExcelWriter('TimeSeries.xlsx', engine='openpyxl')
            category_df.to_excel(excel_file, index=False, sheet_name='TimeSeries Data')
            excel_file._save()
            # Read the file to get bytes
            with open('TimeSeries.xlsx', 'rb') as f:
                excel_data = f.read()
            st.download_button("Download as Excel file", data=excel_data, file_name="TimeSeries.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", help='Click here to download data as Excel file')


    # Create a treemap based on Region,Category,sub-Category
    st.subheader("Hierarchical view of Sales using TreeMap")
    fig3 = px.treemap(filtered_df,path=["Region","Category","Sub-Category"],values="Sales",hover_data=["Sales"],color="Sub-Category")
    fig3.update_layout(width=800,height=650)
    st.plotly_chart(fig3,use_container_width=True)


    chart1,chart2 = st.columns((2))
    with chart1:
        st.subheader('Segment wise Sales')
        fig = px.pie(filtered_df,values = "Sales",names = "Segment", template="plotly_dark")
        fig.update_traces(text = filtered_df["Segment"],textposition = "inside")
        st.plotly_chart(fig,use_container_width=True)

    with chart2:
        st.subheader('Category wise Sales')
        fig = px.pie(filtered_df,values = "Sales",names = "Category", template="gridon")
        fig.update_traces(text = filtered_df["Category"],textposition = "inside")
        st.plotly_chart(fig,use_container_width=True)

    import plotly.figure_factory as ff
    st.subheader(":point_right: Month wise Sub-Category Sales Summary")
    with st.expander("Summary Table"):
        df_sample = df[0:5][["Region","State","City","Category","Sales","Profit","Quantity"]]
        fig = ff.create_table(df_sample,colorscale="Cividis")
        st.plotly_chart(fig,use_container_width=True)

        st.markdown("Month wise sub-category table")
        filtered_df["month"] = filtered_df["Order Date"].dt.month_name()
        sub_category_Year = pd.pivot_table(data=filtered_df,values="Sales",index=["Sub-Category"],columns="month")
        st.write(sub_category_Year)

    # Create a scatter plot
    data1 = px.scatter(filtered_df,x="Sales",y="Profit",size="Quantity")
    data1['layout'].update(title="Relationship between Sales and Profits using Scatter Plot.",titlefont = dict(size=20),xaxis = dict(title="Sales",titlefont=dict(size=19)),yaxis = dict(title="Profit",titlefont=dict(size=19)))

    st.plotly_chart(data1,use_container_width=True)

    with st.expander("View Data"):
        st.write(filtered_df.iloc[:,:])

    # Download original DataSet
    download_type = st.selectbox("Select file format for download:", ['csv', 'xlsx'],key="data_format")
        
    if download_type == 'csv':
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button("Download as CSV file", data=csv, file_name="Data.csv", mime="text/csv", help='Click here to download data as CSV file')

    elif download_type == 'xlsx':
        excel_file = pd.ExcelWriter('Data.xlsx', engine='openpyxl')
        df.to_excel(excel_file, index=False, sheet_name='Data')
        excel_file._save()
        # Read the file to get bytes
        with open('Data.xlsx', 'rb') as f:
            excel_data = f.read()
        st.download_button("Download as Excel file", data=excel_data, file_name="Data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", help='Click here to download data as Excel file')