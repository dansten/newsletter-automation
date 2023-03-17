#Imports
from htmldocx import HtmlToDocx
import pandas as pd
import re
import datetime as dt
import requests
from bs4 import BeautifulSoup
import openai
from stqdm import stqdm
import streamlit as st
from itertools import filterfalse
import math
import os
from dotenv import load_dotenv
load_dotenv()

api_key = os.getenv('API_KEY')

technologies_dict = {'Automation': ['Automation'],
                     'AI': ['AI',  'Artifical Intelligence', 'Machine learning','ML', 'Generative AI', 'Computer vision', 'Natural language', 'NLP', 'NLU', 'Robotics', 'Edge AI', 'Edge', 'Synthetic', 'Synthetic Data', 'Physics-infomred AI', 'Physics Artifical', 'Deep Learning', 'labelling', 'annotation'],
                     'DNA Sequencing': ['DNA', 'DNA sequencing'],
                     'API': ['API', 'Application Programming Interface'],
                     'No-code': ['No-code', 'no code'],
                     'Low-code': ['low-code', 'low code'],
                     'Open Source (OS)': ['open-source', 'open source'],
                     'Graph Technologies': ['graph', 'graph technologies', 'graph technology', 'graph-based'],
                     'IoT': ['IoT', 'internet of things'],
                     'Digital Twins': ['digital twin', 'digital twins', 'digital-twin', 'digital-twins'],
                     '5G/6G': ['5G', '6G'],
                     'Web3 / blockchain': ['web3', 'decentralized','DAO','DeFi','blockchain', 'blockchains', 'crypto', 'cryptocurrency', 'cryptocurrencies', 'NFT', 'NFTs', 'fungible token'],
                     'Metaverse': ['metaverse'],
                     'Autonomous Vehicles': ['autonomous vehicles', 'autonomous cars', 'autonomous vehicle', 'autonomous car', 'autonomous driving','self-driving', 'self driving'],
                     'AR / VR': ['AR', 'VR', 'MR', 'XR', 'virtual reality', 'augmented reality','mixed reality', 'extended reality' ,'virtual'],
                     'Cloud Computing': ['cloud-based', 'cloud', 'cloud computing']}

verticals_list = ['Agriculture', 'Architecture and design', 'Automotive', 'Construction', 'Consumer goods and retail', 'Consulting and legal', 'E-commerce', 'Education', 'Energy and renewable resources', 'Fashion and beauty', 'Financial services', 'Gaming', 'Government', 'Healthcare', 'Hospitality and tourism', 'Information technology', 'Insurance', 'Life sciences & pharma', 'Media & entertainment', 'Metals & mining', 'Non-profit and charity', 'Oil & gas', 'Manufacturing', 'Real estate', 'Sports and fitness', 'Telco', 'Transportation and logistics', 'Utilities', 'E-commerce', 'Advertisement & Marketing', 'Non-vertical specific']
verticals = ", ".join(verticals_list)

pattern = re.compile(r'\((?:[^)(]+|\((?:[^)(]+|\([^)(]*\))*\))*\)')

def fix_url(url):
  if type(url) == float:
    url = ""
  elif url.startswith('www'):
    url = 'https://'+url
  elif url.startswith('https://'):
    url = url
  else: 
    url = "https://www." + url
  return url

def remove_parentheses(text):
    cleaned_text = re.sub(pattern, '', text).rstrip()
    return cleaned_text

def pre_process(data):
    try:
        data = data.dropna(subset=['Year Founded'])
        data['Employees'] = data['Employees'].fillna(0)
        data['SimilarWeb Unique Visitors'] = data['SimilarWeb Unique Visitors'].str.replace(',', '').astype(float)
        data['Total Raised'] = data['Total Raised'].astype(float)
        data['Last Financing Size'] = data['Last Financing Size'].astype(float)
        data['Employees'] = data['Employees'].astype(float)
        data["Website"] = data["Website"].apply(lambda x: fix_url(x))
        data = data[data['Companies'].notna()]
        data['Companies'] = data['Companies'].apply(remove_parentheses)
        data['Year Founded'] = data['Year Founded'].astype(int)
    except:
        print('There is an error with your input data...')
    return data


def check_technologies(description, technologies_dict):
    technologies = []
    for tech, tech_values in technologies_dict.items():
        for val in tech_values:
            if re.search(rf"\b{val}\b", description, re.IGNORECASE):
                technologies.append(tech)
                break
    return technologies

def get_website_text(url):
    headers = {"user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.182 Safari/537.36"}
    try:
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        text = soup.get_text()
        clean_text = " ".join(text.split())
        return clean_text
    except requests.exceptions.RequestException as e:
        print(f"Error retrieving website text: {e}")
        return ""

def get_full_description(df, scrape=False):
    descriptions = []
    for _, row in df.iterrows():
        desc = row['Description']
        if scrape:
            website_text = get_website_text(row['Website'])
            desc = f'Company Description: {desc}. \n Also use the following information from their website: {website_text}'
        if len(desc) > 4100:
            desc = desc[:4100]
        descriptions.append(desc)
    return descriptions


def get_short_descriptions(data, api_key):
    openai.api_key = api_key
    short_descs = []
    for index, desc in enumerate(data['Full Description']):
        company = data['Companies'].iloc[index]
        prompt = f'Hi, can you please explain what {company} does in a clear an precise manner? Restrain yourself to 25 words. {desc}'
        try:
            completion = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                temperature = 0.2,
                messages=[
                {"role": "user", "content": prompt}
                ]
            )
            short_descs.append(completion.choices[0].message.content.strip('\n.'))
        except:
            print("Error with ChatGPT...")
            short_descs.append("")
    return short_descs


def get_verticals(data, api_key):
    openai.api_key = api_key
    vertical_list = []
    progress_text = "Operation in progress. Please wait."
    my_bar = st.progress(0, text=progress_text)
    data_len = data.shape[0]
    for index, desc in stqdm(enumerate(data['Full Description'])):
        
        company = data['Companies'].iloc[index]
        prompt = f'Which vertical does {company} target out of the following: {verticals}. Please provide your answer 1-2 words. If unsure, simply answer \'Other\'. Description of {company}: {desc}'
        
        try:
            completion = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            temperature = 0,
            messages=[
                {"role": "user", "content": prompt}
            ]
            )
            vertical_list.append(completion.choices[0].message.content.strip('\n.'))
        except:
            vertical_list.append("")
            print("Error...")
        my_bar.progress((index + 2) * math.floor(100/data_len), text=progress_text)
    return vertical_list

def update_company_info(data, api_key):
    data['Full Description'] = get_full_description(data[['Description', 'Website']], scrape=True)
    data['Short Description'] = get_short_descriptions(data, api_key)
    data['Verticals'] = get_verticals(data, api_key)
    return data

def format_employee_info(row):
    employee_history = str(row.get('Employee History', ''))
    
    if not isinstance(employee_history, str):
        return ""
    
    employees = []
    employee_years = []
    
    for year, count in re.findall(r'(\d+):\s*([\d,]+)', employee_history):
        employees.append(int(count.replace(',', '')))
        employee_years.append(int(year))
    
    momentum_str = ""
    if employees:
        ftes = employees[-1]
        ftes_change = ((ftes - employees[0]) / employees[0]) * 100 if len(employees) > 1 else 0
        ftes_change_str = f"{ftes_change:+.0f}%" if ftes_change != 0 else "0%"
        momentum_str = f"<li>FTEs: <b>{ftes} / {ftes_change_str} YoY</b></li>"
    
    return momentum_str

def format_traffic_info(row):
    visitors = row.get('SimilarWeb Unique Visitors', '')
    visitors_change = row.get('SimilarWeb Unique Visitors % Change', '')
    if not visitors or not visitors_change:
        return ''
    if visitors == 0:
        return ''
    visitors_str = str(visitors)
    if visitors >= 1000:
        visitors_str = f'{visitors/1000:,.1f}k'
    visitors_change = float(str(visitors_change).strip('%')) / 100.0
    traffic_str = f"<li>Traffic: <b>{visitors_str} / {visitors_change:+.0%} MoM</b></li>"
    return traffic_str

def format_company(row, scrape=filterfalse):
    company_name = row['Companies']
    url = row['Website']    
    description = row['Short Description']

    tech_tags = row['Technologies']
    verticals = row['Verticals']

    year_founded = row['Year Founded']
    hq_location = row['HQ Location']
    #first_sentence = re.match(r'^(.+?[\.\?\!])', description).group(1) if (re.match(r'^(.+?[\.\?\!])', description)) else row['Short Description']

    total_funding = row['Total Raised']
    if pd.isna(total_funding):
        total_funding = 'No funding to date'
    else:
        total_funding = f'€{round(total_funding, 1)}m'

    investors = row['Active Investors']
    last_funding_amount = row['Last Financing Size']
    if not pd.isna(last_funding_amount):
        last_funding_amount = f'€{round(last_funding_amount, 1)}m'

    last_funding_date = ''
    if not pd.isna(row['Last Financing Date']):
        last_funding_date = dt.datetime.strptime(row['Last Financing Date'], '%d-%b-%Y').strftime('%b-%y')

    employee_info_str = format_employee_info(row)
    traffic_info_str = format_traffic_info(row)
    
    if not row['SimilarWeb Unique Visitors'] or row['SimilarWeb Unique Visitors'] == 0:
        momentum_str = ''
    else:
        momentum_str = f'<li>Momentum:<ul>{employee_info_str}{traffic_info_str}</ul></li>'

    if len(tech_tags)==0:
      tech_tags_str = ''
    else:
      tech_tags_str = f'<p><b>Technology tags</b>: {", ".join(tech_tags)}</p>' 
    return (
        f'<p><b><u><a href="{url}">{company_name}</a></u></b> ({hq_location}, {year_founded}): {description}</p>'
        f'<p><b>Vertical Tags</b>: {verticals}</p>'
        f'{tech_tags_str}'
        f'<ul><li>Funding: <b>{total_funding}</b> from <b>{investors}</b><ul><li>Last round: <b>{last_funding_amount}</b> in <b>{last_funding_date}</b></li></ul></li>'
        f'{momentum_str}</ul>'
    )

def main():
    st.markdown("# SaaS Newsletter Automation")
    st.markdown("Instructions:")
    st.markdown("- Go to Pitchbook and download the data for the relevant companies.")
    st.markdown("- Make sure your Pitchbook export has the following columns: **Companies** | **Website** | **HQ Location** | **Year Founded** | **Total Raised** | **Last Financing Size** | **Last Financing Date** | **Description** | **Active Investors** | **Employees** | **Employee History** | **SimilarWeb Unique Visitors** | **SimilarWeb Unique Visitors % Change**")
    st.markdown("- In the export, remove all the rows before the header columns, as well as remove the Pitchbook image.")
    st.markdown("- Save the file as a CSV.")
    st.markdown("- Upload the CSV file using the upload button below.")
    st.markdown("- Wait for the program to finish executing (this may take 1-2 minutes).")
    st.markdown("If you get an error, there is most likley something wrong with the input data.")

    # File upload through Streamlit
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

    if uploaded_file is not None:
        file_name = uploaded_file.name[:-4] # Remove .csv extension
        data = pd.read_csv(uploaded_file)
        if st.button("Generate Word Document", type="secondary"):
            data = pre_process(data)
            data["Technologies"] = data["Description"].apply(lambda x: check_technologies(x, technologies_dict))
            data = update_company_info(data, api_key)

            
            htmlstring = ""

            for index, row in data.iterrows():
                company_link = format_company(row)
                htmlstring = htmlstring + company_link

            with open(f'{file_name}.html', "w") as file:
                file.write(htmlstring)

            new_parser = HtmlToDocx()
            new_parser.paragraph_style = 'No Spacing'
            new_parser.parse_html_file(f'{file_name}.html', f'{file_name}')

            with open(f'{file_name}.docx', 'rb') as doc:
                doc_bytes = doc.read()
                st.download_button(label='Download Word Document', data=doc_bytes, file_name=f'{file_name}.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
main()