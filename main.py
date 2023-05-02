import pandas as pd
import snowflake.connector
import streamlit as st
from st_aggrid import AgGrid
from st_aggrid.grid_options_builder import GridOptionsBuilder
from st_aggrid import GridUpdateMode, DataReturnMode
import warnings
from yahooquery import Ticker
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
from datetime import date
from PIL import Image
import requests
import os
from io import BytesIO
import openai


# hide future warnings (caused by st_aggrid)
warnings.simplefilter(action='ignore', category=FutureWarning)

#set page layout and define basic variables
st.set_page_config(layout="wide", page_icon='⚡', page_title="Instant Insight")
path = os.path.dirname(__file__)
today = date.today()


# get Snowflake credentials from Streamlit secrets
SNOWFLAKE_ACCOUNT = st.secrets["snowflake_credentials"]["SNOWFLAKE_ACCOUNT"]
SNOWFLAKE_USER = st.secrets["snowflake_credentials"]["SNOWFLAKE_USER"]
SNOWFLAKE_PASSWORD = st.secrets["snowflake_credentials"]["SNOWFLAKE_PASSWORD"]
SNOWFLAKE_DATABASE = st.secrets["snowflake_credentials"]["SNOWFLAKE_DATABASE"]
SNOWFLAKE_SCHEMA = st.secrets["snowflake_credentials"]["SNOWFLAKE_SCHEMA"]


@st.cache_resource
def get_database_session():
    """Returns a database session object."""
    return snowflake.connector.connect(
        account=SNOWFLAKE_ACCOUNT,
        user=SNOWFLAKE_USER,
        password=SNOWFLAKE_PASSWORD,
        database=SNOWFLAKE_DATABASE,
        schema=SNOWFLAKE_SCHEMA,
    )


@st.cache_data
def get_data(_conn):
    """Returns a pandas DataFrame with the data from Snowflake."""
    query = 'SELECT * FROM us_prospects;'
    cur = conn.cursor()
    cur.execute(query)

    # Fetch the result as a pandas DataFrame
    column_names = [col[0] for col in cur.description]
    data = cur.fetchall()
    df = pd.DataFrame(data, columns=column_names)

    # Close the connection to Snowflake
    cur.close()
    conn.close()

    return df


def resize_image(url):
    """function to resize logos while keeping aspect ratio. Accepts URL as an argument and return an image object"""

    # Open the image file
    image = Image.open(requests.get(url, stream=True).raw)

    # if a logo is too high or too wide then make the background container twice as big
    if image.height > 140:
        container_width = 220 * 2
        container_height = 140 * 2

    elif image.width > 220:
        container_width = 220 * 2
        container_height = 140 * 2
    else:
        container_width = 220
        container_height = 140

    # Create a new image with the same aspect ratio as the original image
    new_image = Image.new('RGBA', (container_width, container_height))

    # Calculate the position to paste the image so that it is centered
    x = (container_width - image.width) // 2
    y = (container_height - image.height) // 2

    # Paste the image onto the new image
    new_image.paste(image, (x, y))

    return new_image


def add_image(slide, image, left, top, width):
    """function to add an image to the PowerPoint slide and specify its position and width"""
    slide.shapes.add_picture(image, left=left, top=top, width=width)


# function to replace text in pptx first slide with selected filters
def replace_text(replacements, shapes):
    """function to replace text on a PowerPoint slide. Takes dict of {match: replacement, ... } and replaces all matches"""
    for shape in shapes:
        for match, replacement in replacements.items():
            if shape.has_text_frame:
                if (shape.text.find(match)) != -1:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        whole_text = "".join(run.text for run in paragraph.runs)
                        whole_text = whole_text.replace(str(match), str(replacement))
                        for idx, run in enumerate(paragraph.runs):
                            if idx != 0:
                                p = paragraph._p
                                p.remove(run._r)
                        if bool(paragraph.runs):
                            paragraph.runs[0].text = whole_text



def get_stock(ticker, period, interval):
    """function to get stock data from Yahoo Finance. Takes ticker, period and interval as arguments and returns a DataFrame"""
    hist = ticker.history(period=period, interval=interval)
    hist = hist.reset_index()
    # capitalize column names
    hist.columns = [x.capitalize() for x in hist.columns]

    return hist


def plot_graph(df, x, y, title, name):
    """function to plot a line graph. Takes DataFrame, x and y axis, title and name as arguments and returns a Plotly figure"""
    fig = px.line(df, x=x, y=y, template='simple_white',
                        title='<b>{} {}</b>'.format(name, title))
    fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')

    return fig


def get_financials(df, col_name, metric_name):
    """function to get financial metrics from a DataFrame. Takes DataFrame, column name and metric name as arguments and returns a DataFrame"""
    metric = df.loc[:, ['asOfDate', col_name]]
    metric_df = pd.DataFrame(metric).reset_index()
    metric_df.columns = ['Symbol', 'Year', metric_name]

    return metric_df


def generate_gpt_response(gpt_input, max_tokens):
    """function to generate a response from GPT-3. Takes input and max tokens as arguments and returns a response"""
    # load api key from secrets
    openai.api_key = st.secrets["openai_credentials"]["API_KEY"]

    completion = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        max_tokens=max_tokens,
        temperature=0,
        messages=[
            {"role": "user", "content": gpt_input},
        ]
    )
    gpt_response = completion.choices[0].message['content'].strip()
    return gpt_response


# Get the data from Snowflake
conn = get_database_session()
df = get_data(conn)

# select columns to show
df_filtered = df[['COMPANY_NAME', 'SECTOR', 'INDUSTRY', 'PROSPECT_STATUS', 'PRODUCT']]

#create sidebar filters
st.sidebar.write('**Use filters to select prospects** 👇')
unique_sector = sorted(df['SECTOR'].unique())
sector_checkbox = st.sidebar.checkbox('All Sectors', help='Check this box to select all sectors')

#if select all checkbox is checked then select all sectors
if sector_checkbox:
    selected_sector = st.sidebar.multiselect('Select Sector', unique_sector, unique_sector)
else:
    selected_sector = st.sidebar.multiselect('Select Sector', unique_sector)

#if a user selected sector then allow to check all industries checkbox
if len(selected_sector) > 0:
    industry_checkbox = st.sidebar.checkbox('All Industries', help='Check this box to select all industries')
    # filtering data
    df_filtered = df_filtered[(df_filtered['SECTOR'].isin(selected_sector))]
    # show number of selected customers
    num_of_cust = str(df_filtered.shape[0])
else:
    industry_checkbox = st.sidebar.checkbox('All Industries', help='Check this box to select all industries',
                                           disabled=True)
    # show number of selected customers
    num_of_cust = str(df_filtered.shape[0])
    df_filtered = df_filtered[['COMPANY_NAME', 'SECTOR', 'INDUSTRY', 'PROSPECT_STATUS', 'PRODUCT']]

#if select all checkbox is checked then select all industries
unique_industry = sorted(df['INDUSTRY'].loc[df['SECTOR'].isin(selected_sector)].unique())
if industry_checkbox:
    selected_industry = st.sidebar.multiselect('Select Industry', unique_industry, unique_industry)
else:
    selected_industry = st.sidebar.multiselect('Select Industry', unique_industry)

#if a user selected industry then allow them to check all statuses checkbox
if len(selected_industry) > 0:
    status_checkbox = st.sidebar.checkbox('All Prospect Statuses', help='Check this box to select all prospect statuses')
    # filtering data
    df_filtered = df_filtered[(df_filtered['SECTOR'].isin(selected_sector)) & (df_filtered['INDUSTRY'].isin(selected_industry))]
    # show number of selected customers
    num_of_cust = str(df_filtered.shape[0])

else:
    status_checkbox = st.sidebar.checkbox('All Prospect Statuses', help='Check this box to select all prospect statuses', disabled=True)

unique_status = sorted(df_filtered['PROSPECT_STATUS'].loc[df_filtered['SECTOR'].isin(selected_sector) & df_filtered['INDUSTRY'].isin(selected_industry)].unique())

#if select all checkbox is checked then select all statuses
if status_checkbox:
    selected_status = st.sidebar.multiselect('Select Prospect Status', unique_status, unique_status)
else:
    selected_status = st.sidebar.multiselect('Select Prospect Status', unique_status)


#if a user selected status then allow them to check all products checkbox
if len(selected_status) > 0:
    product_checkbox = st.sidebar.checkbox('All Products', help='Check this box to select all products')
    # filtering data
    df_filtered = df_filtered[(df_filtered['SECTOR'].isin(selected_sector)) & (df_filtered['INDUSTRY'].isin(selected_industry)) & (df_filtered['PROSPECT_STATUS'].isin(selected_status))]
    # show number of selected customers
    num_of_cust = str(df_filtered.shape[0])

else:
    product_checkbox = st.sidebar.checkbox('All Products', help='Check this box to select all products', disabled=True)

unique_products = sorted(df_filtered['PRODUCT'].loc[df_filtered['SECTOR'].isin(selected_sector) &
                                                    df_filtered['INDUSTRY'].isin(selected_industry)
                                                    & df_filtered['PROSPECT_STATUS'].isin(selected_status)].unique())

#if select all checkbox is checked then select all products
if product_checkbox:
    selected_product = st.sidebar.multiselect('Select Product', unique_products, unique_products)
else:
    selected_product = st.sidebar.multiselect('Select Product', unique_products)

if selected_product:
    # filtering data
    df_filtered = df_filtered[(df_filtered['SECTOR'].isin(selected_sector)) & (df_filtered['INDUSTRY'].isin(selected_industry))
                              & (df_filtered['PROSPECT_STATUS'].isin(selected_status)) & (df_filtered['PRODUCT'].isin(selected_product))]
    # show number of selected customers
    num_of_cust = str(df_filtered.shape[0])


with st.sidebar:
    st.markdown('''The dataset is taken from [Kaggle](https://www.kaggle.com/datasets/aramacus/usa-public-companies) and slightly modified for the purpose of this app.
    ''', unsafe_allow_html=True)
    st.markdown('''[GitHub Repo](https://github.com/arsentievalex/instant-insight-web-app)''', unsafe_allow_html=True)
    st.markdown('''The app created by [Oleksandr Arsentiev](https://twitter.com/alexarsentiev) for the purpose of
    Streamlit Summit Hackathon''', unsafe_allow_html=True)

##############################################################################################################

st.title('Welcome to the Instant Insight App!⚡')

with st.expander('What is this app about?'):
    st.write('''
    This app is designed to generate an instant company research.\n
    In a matter of few clicks, a user gets a PowerPoint presentation with the company overview, SWOT analysis, financials, and value propostion tailored for the selling product. 
    The app works with the US public companies.
    
    Use Case Example:\n
    Imagine working in sales for a B2B SaaS company that has hundreds of prospects and offers the following products: 
    Accounting and Planning Software, CRM, Chatbot, and Cloud Data Storage.
    You are tasked to do a basic prospect research and create presentations for your team. The prospects data is stored in a Snowflake database that feeds your CRM system.
    You can use this app to quickly filter the prospects by sector, industry, prospect status, and product. 
    Next, you can select the prospect you want to include in the presentation and click the button to generate the presentation.
    And...that's it! You have the slides ready to be shared with your team.
    
    Tech Stack:\n
    • Database - Snowflake via Snowflake Connector\n
    • Data Processing - Pandas\n
    • Research Data - Yahoo Finance via Yahooquery, GPT 3.5 via OpenAI API\n
    • Visualization - Plotly\n
    • Frontend - Streamlit, AgGrid\n
    • Presentation - Python-pptx\n
    ''')


st.metric(label='Number of Prospects', value=num_of_cust)

# button to create slides
ui_container = st.container()
with ui_container:
    submit = st.button(label='Generate Presentation')

# apply proper capitalization to column names and replace underscore with space
df_filtered.columns = df_filtered.columns.str.title().str.replace('_', ' ')

# creating AgGrid dynamic table and setting configurations
gb = GridOptionsBuilder.from_dataframe(df_filtered)
gb.configure_selection(selection_mode="single", use_checkbox=True)
gb.configure_column(field='Company Name', width=270)
gb.configure_column(field='Sector', width=260)
gb.configure_column(field='Industry', width=350)
gb.configure_column(field='Prospect Status', width=270)
gb.configure_column(field='Product', width=250)

gridOptions = gb.build()

response = AgGrid(
    df_filtered,
    gridOptions=gridOptions,
    enable_enterprise_modules=False,
    height=600,
    update_mode=GridUpdateMode.SELECTION_CHANGED,
    data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
    fit_columns_on_grid_load=False,
    theme='alpine',
    allow_unsafe_jscode=True
)

# get selected rows
response_df = pd.DataFrame(response["selected_rows"])

# if user input is empty and button is clicked then show warning
if submit and response_df.empty:
    with ui_container:
        st.warning("Please select a prospect!")

# if user input is not empty and button is clicked then generate slides
elif submit and response_df is not None:
    with ui_container:
        with st.spinner('Generating awesome slides for you...⏳'):

            try:
                # define variables for selected prospect
                company_name = response_df['Company Name'].values[0]
                product = response_df['Product'].values[0]

                # join df with response_df to get a ticker of selected prospect
                df_ticker = pd.merge(df, response_df, left_on='COMPANY_NAME', right_on='Company Name')
                selected_ticker = df_ticker['TICKERS'].values[0]

                # open presentation template
                pptx = path + '//' + 'template.pptx'
                prs = Presentation(pptx)

                ticker = Ticker(selected_ticker)

                # get stock info
                name = ticker.price[selected_ticker]['shortName']
                sector = ticker.summary_profile[selected_ticker]['sector']
                industry = ticker.summary_profile[selected_ticker]['industry']
                employees = ticker.summary_profile[selected_ticker]['fullTimeEmployees']
                country = ticker.summary_profile[selected_ticker]['country']
                city = ticker.summary_profile[selected_ticker]['city']
                website = ticker.summary_profile[selected_ticker]['website']
                summary = ticker.summary_profile[selected_ticker]['longBusinessSummary']
                logo_url = 'https://logo.clearbit.com/' + website

                # declare pptx variables
                title_slide = prs.slides[0]
                summary_slide = prs.slides[1]
                s_w_slide = prs.slides[2]
                strategy_slide = prs.slides[4]

                # initialize a list of shapes
                shapes_1 = []
                shapes_2 = []
                shapes_3 = []
                shapes_4 = []

                # create lists with shape objects
                for shape in title_slide.shapes:
                    shapes_1.append(shape)

                for shape in summary_slide.shapes:
                    shapes_2.append(shape)

                for shape in s_w_slide.shapes:
                    shapes_3.append(shape)

                for shape in strategy_slide.shapes:
                    shapes_4.append(shape)

                # initiate a dictionary of placeholders and values to replace
                replaces_1 = {
                    '{company}': name,
                    '{date}': today}

                replaces_2 = {
                    '{c}': name,
                    '{s}': sector,
                    '{i}': industry,
                    '{co}': country,
                    '{ci}': city,
                    '{ee}': "{:,}".format(employees),
                    '{w}': website,
                    '{summary}': summary
                }

                # run the function to replace placeholders with values
                replace_text(replaces_1, shapes_1)
                replace_text(replaces_2, shapes_2)

                # check if a logo ulr returns code 200 (working link)
                if requests.get(logo_url).status_code == 200:
                    #create logo image object
                    logo = resize_image(logo_url)
                    logo.save('logo.png')
                    logo_im = 'logo.png'

                    # add logo to the slide
                    add_image(prs.slides[1], image=logo_im, left=Inches(1.2), width=Inches(2), top=Inches(0.5))
                    os.remove('logo.png')

                ##############################################################################################################
                # create slides with financial plots

                # get financial data
                income_df = ticker.income_statement()
                valuation_df = ticker.valuation_measures

                # plot stock price
                stock_df = get_stock(ticker=ticker, period='5y', interval='1mo')
                stock_fig = plot_graph(df=stock_df, x='Date', y='Open', title='Stock Price USD', name=name)

                stock_fig.write_image("stock.png")
                stock_im = 'stock.png'

                add_image(prs.slides[3], image=stock_im, left=Inches(1.8), width=Inches(4.5), top=Inches(0.5))
                os.remove('stock.png')

                # plot revenue
                rev_df = get_financials(df=income_df, col_name='TotalRevenue', metric_name='Total Revenue')
                rev_fig = plot_graph(df=rev_df, x='Year', y='Total Revenue', title='Total Revenue USD', name=name)

                rev_fig.write_image("rev.png")
                rev_im = 'rev.png'

                add_image(prs.slides[3], image=rev_im, left=Inches(1.8), width=Inches(4.5), top=Inches(3.8))
                os.remove('rev.png')

                # plot market cap
                marketcap_df = get_financials(df=valuation_df, col_name='MarketCap', metric_name='Market Cap')
                marketcap_fig = plot_graph(df=marketcap_df, x='Year', y='Market Cap', title='Market Cap USD', name=name)

                marketcap_fig.write_image("marketcap.png")
                marketcap_im = 'marketcap.png'

                add_image(prs.slides[3], image=marketcap_im, left=Inches(7.3), width=Inches(4.5), top=Inches(0.5))
                os.remove('marketcap.png')

                # plot ebitda
                ebitda_df = get_financials(df=income_df, col_name='NormalizedEBITDA', metric_name='EBITDA')
                ebitda_fig = plot_graph(df=ebitda_df, x='Year', y='EBITDA', title='EBITDA USD', name=name)

                ebitda_fig.write_image("ebitda.png")
                ebitda_im = 'ebitda.png'

                add_image(prs.slides[3], image=ebitda_im, left=Inches(7.3), width=Inches(4.5), top=Inches(3.8))
                os.remove('ebitda.png')

                ############################################################################################################
                # create strengths and weaknesses slide

                input_swot = 'What are the strengths, weaknesses, opportunities and threats of {} company with ticker {}? ' \
                            'Provide only strengths, weaknesses, opportunities and threats without any extra text in the given order, be specific and consise, ' \
                        'use bullet points.'.format(name, selected_ticker)

                # return response from GPT-3
                gpt_swot = generate_gpt_response(input_swot, 1500)
                swot_title = 'SWOT Analysis of {}'.format(name)

                # initiate a dictionary of placeholders and values to replace
                replaces_3 = {
                    '{swot}': gpt_swot.strip(),
                    '{swot_title}': swot_title}

                # run the function to replace placeholders with values
                replace_text(replaces_3, shapes_3)

                ############################################################################################

                # create value prop slide
                input_value = 'What needs and pain points {name} company may have that can be solved with a {product}? ' \
                              '{name} is a US public company with ticker {ticker} that operates in {industry} industry. ' \
                              'Create a brief value proposition taking into account what the company does and its industry. Be specific and use bullet points'.\
                    format(name=name, product=product, ticker=selected_ticker, industry=industry)

                # return response from GPT-3
                gpt_value = generate_gpt_response(input_value, 1500)

                vp_title = 'Value Proposition of {} for {}'.format(product, name)

                # initiate a dictionary of placeholders and values to replace
                replaces_4 = {
                    '{vp}': gpt_value.strip(),
                    '{vp_title}': vp_title}

                # run the function to replace placeholders with values
                replace_text(replaces_4, shapes_4)

                ############################################################################################

                # create file name
                filename = '{} {}.pptx'.format(name, today)

                # save presentation as binary output
                binary_output = BytesIO()
                prs.save(binary_output)

                # display success message and download button
                with ui_container:
                    st.success('The slides have been generated! :tada:')

                    st.download_button(label='Click to download PowerPoint',
                                       data=binary_output.getvalue(),
                                       file_name=filename)

            # if there is any error, display an error message
            except Exception as e:
                with ui_container:
                    st.error("Oops, something went wrong, please try again or select a different prospect.")
