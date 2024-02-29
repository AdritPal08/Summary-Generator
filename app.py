import sys
from exception import CustomException
from logger import logging
from dotenv import load_dotenv

import google.generativeai as genai
import textwrap
import pandas as pd
import streamlit as st
import io
import os

# take environment variables from .env.
env_loaded = load_dotenv()
if env_loaded:
    logging.info("Environment variables loaded.")
else:
    logging.error("Environment variables not loaded.")

# api_key = os.getenv("GOOGLE_API_KEY")
api_key = st.secrets["GOOGLE_API_KEY"]
if api_key:
    genai.configure(api_key=api_key)
else:
    logging.error("API key not found or empty.")
    
    

def to_markdown(text):
  # Remove the indentation from the text
  text = textwrap.dedent(text)
  # Add the markdown syntax manually
  # text = "\n" + text.replace("•", "*")
  text = text.replace("•", " ")
  text = text.replace("*", " ")
  text = text.replace("-", " ")
  # Return the markdown-formatted string
  return text

# Define a function to convert the DataFrame to an Excel file in memory
def to_excel(df):
  # Create a BytesIO object
  output = io.BytesIO()
  # Create an ExcelWriter object with the output as the file
  writer = pd.ExcelWriter(output, engine="xlsxwriter")
  # Write the DataFrame to the ExcelWriter object
  df.to_excel(writer, index=False, sheet_name="Summary")
  # Save and close the ExcelWriter object
  writer.close() # Use the close() method instead of the save() method
  # Return the output as bytes
  return output.getvalue()


generation_config = {
  "temperature": 0.4,
  "top_p": 1,
  "top_k": 32,
  "max_output_tokens": 4096,
}


model = genai.GenerativeModel(model_name="gemini-pro", generation_config=generation_config)

# initialize our streamlit app
st.set_page_config(page_title="CompanyInsight Pro")
st.header("CompanyInsight Pro")
st.write("An app that generates short summaries from given company names.")
logging.info("Streamlit app initialized.")

# Create a sidebar
sidebar = st.sidebar

# Get the excel file
uploaded_file = st.sidebar.file_uploader("Upload your excel file: ", type=["xlsx"])


# Get the text prompt from the user
company_input = st.sidebar.text_input("Enter Company Name (if providing a single name): ", key="input")
logging.info(f"User input comapny: {company_input}")

# Check if the file is not None
if uploaded_file is not None:
    # Read the excel file as a dictionary of DataFrames
    df = pd.read_excel(uploaded_file, sheet_name=None)
    # Get the sheet names as a list of strings
    tabs = list(df.keys())
    # Get the sheet name from the user
    sheet_name = st.sidebar.selectbox("Choose sheet name ", options=tabs, index=0)
    # Get the DataFrame corresponding to the selected sheet
    data = df[sheet_name]
    # Get the sheet names as a list of strings
    columns = list(data.columns)
    # Get the column name from the user
    column_name = st.sidebar.selectbox("Choose column name ", options=columns, index=0)
    companies = data[column_name]
    # # Display the DataFrame
    # st.dataframe(data[column_name])
elif company_input !="":
  st.write("")
else:
    # Display a message if no file is uploaded
    st.write("Please upload an Excel file or enter a company name to proceed.")


submit=st.sidebar.button("⌛Summarize Now..", key="submit")

company_name = []
summary_company = []
if submit:
  if uploaded_file is not None and company_input == "":
    
    try:
        # Create a spinner object
        with st.spinner('Generating response...'):
            st.subheader("The Summaries are:")
            # Create a progress bar object
            progress_bar = st.progress(0)
            # Calculate the increment value for the progress bar
            increment = 1.0 / len(companies)
            # Loop through the companies
            for i, company in enumerate(companies):
                # prompt = f"Write a short summary of {company} in a formal tone. Summary will be minimum of 5 lines to maximum of 10 lines." 
                prompt = f'''Write a summary of the company:{company} in a formal tone and with as much detail as possible. Don't add any wrong information. All the information should be correct.  
                Please summarize the company with all the relevant information within a minimum of 8 lines and a maximum of 10 lines.'''   
                response = model.generate_content(prompt)
                summary = to_markdown(response.text)
                company_name.append(company)
                summary_company.append(summary)
                # Use st.write to display the company name
                st.write(f"**{company}**")
                # Use st.markdown to display the summary content
                st.markdown(summary)
                # Update the progress bar
                progress_bar.progress((i + 1) * increment)
    except Exception as e:
        logging.error(e)
        st.error(f"An error occurred: {e}")
  elif company_input != "":
    with st.spinner('Generating response...'):
      st.subheader("The Summaries are:")
      # prompt = f"Write a short summary of {company_input} in a formal tone. Summary will be minimum of 5 lines to maximum of 10 lines."
      prompt = f'''Write a summary of the company:{company_input} in a formal tone and with as much detail as possible. Don't add any wrong information. All the information should be correct.  
                Please summarize the company with all the relevant information within a minimum of 8 lines and a maximum of 10 lines.'''
      response = model.generate_content(prompt)
      summary = to_markdown(response.text)
      company_name.append(company_input)
      summary_company.append(summary)
      # Use st.write to display the company name
      st.write(f"**{company_input}**")
      # Use st.markdown to display the summary content
      st.markdown(summary)
  else:
    st.write("Please enter a text prompt to proceed.")
    st.stop() # Stop the execution of the app
Summary_data = pd.DataFrame(list(zip(company_name, summary_company)), columns=["Company Name", "Summary"])
excel_data = to_excel(Summary_data)

# Create a download button with the excel_data as the data argument
st.sidebar.download_button(
  label="Download Excel file",
  data=excel_data,
  file_name="summary_data.xlsx",
  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Add a footer with your name
footer_css = st.sidebar.markdown("""
<style>
.footer {
  position: fixed;
  bottom: 0;
  right: 0;
  padding: 10px;
  color: white;
  font-size: 12px;
}
</style>
""", unsafe_allow_html=True)

footer = st.sidebar.markdown("""
<div class="footer">
Created by : Adrit Pal
</div>
""", unsafe_allow_html=True)