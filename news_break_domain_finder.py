from googlesearch import search
import requests
from bs4 import BeautifulSoup
import pandas as pd
from serpapi import GoogleSearch
import os
from datetime import datetime

def save_result_sheet(google_search_result, sheet_name="Result10x"):

    # Convert the dictionary to DataFrame
    results_df = pd.DataFrame(google_search_result)

    # Shift the index to start from 1
    results_df.index = results_df.index + 1

    # Check if the file exists
    if os.path.exists(excel_file_path):
        with pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            try:
                # Load the existing sheet
                existing_data = pd.read_excel(excel_file_path, sheet_name=sheet_name)
                # Append new data to the existing data
                updated_data = pd.concat([existing_data, results_df], ignore_index=True)
            except ValueError:
                # If the sheet does not exist, just use the new data
                updated_data = results_df

            # Write the updated data back to the specified sheet
            updated_data.to_excel(writer, sheet_name=sheet_name, index=False)


    else:
        # Save results to a new sheet in the same Excel file
        with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
            results_df.to_excel(writer, sheet_name=sheet_name, index=True)

    print(f"Results saved to sheet {sheet_name} in the file: {excel_file_path}")
    print("""
            -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
""")

# to search

url_lst =[]

def google_search(domain_list,num_results=3):

    results = []
    
    for domain_name in domain_list:
        query = f"""
                    site:newsbreak.com "{domain_name}"
        """
        url_lst =[]

        print(f"Searching for: {query}")

        try:
            # Perform the Google search
            for search_result_url in search(query, lang="en", sleep_interval=1, num_results=num_results, region="us", advanced=False):
                print(search_result_url)
                url_lst.append(search_result_url)

        except requests.exceptions.RequestException as e:
            print(f"Error fetching {domain_name}: {e}")
            return False

        for url_check in url_lst:

            try:
                print(f"Scraping URL: {url_check}")
                response = requests.get(url_check)
                response.raise_for_status()  # Check for HTTP errors
                
                #Parse the HTML content
                soup = BeautifulSoup(response.text, 'html.parser')
                page_text = soup.get_text()  # Extract text from the page

                # Check if the domain is in the page text
                print(f"Checking if '{domain_name}' is in the page text...{url_check}")

                if domain_name in page_text:
                    result =  {
                        "Domain Name": domain_name,
                        "URL": url_check,
                        "Status": True
                    }

                    break

                else:
                    result = {
                        "Domain Name": domain_name,
                        "URL": url_check,
                        "Status": False
                    }

                results.append(result)

                        
            except requests.exceptions.RequestException as e:
                print(f"Error scraping in {url_check}: {e}")
                False
            
    return results


    print(url_lst)

    return(url_lst)

def check_domain_in_url(url, domain_name):
    try:
        print(f"Scraping URL: {url}",end="\n")
        response = requests.get(url)
        response.raise_for_status()  # Check for HTTP errors
        
        #Parse the HTML content
        soup = BeautifulSoup(response.text, 'html.parser')
        page_text = soup.get_text()  # Extract text from the page

        # Check if the domain is in the page text
        print(f"Checking if '{domain_name}' is in the page text...{url}",end="\n")

        if domain_name in page_text:
            return True
        else:
            return False
                        
    except requests.exceptions.RequestException as e:
        print(f"Error scraping in {url}: {e}",end="\n")
        False
        

def serp_google_search(domain_list):
    

    for domain_name in domain_list:

        Domain_name = []
        Result_states = []
        URLs = []
        Status = []

        google_search_result = {}

        params = {
        "api_key": "25190da1d388dd276fb7427f203916eb05102dd95600eb063e15492d291fede6",
        "engine": "google",
        "q": f"""
            site:newsbreak.com "{domain_name}"
        """,
        "location": "United States",
        "google_domain": "google.com",
        "gl": "us",
        "hl": "en"
        }

        print(f"Searching for: {params['q']}",end="\n")

        search = GoogleSearch(params)
        search_results = search.get_dict()

        # Print the organic_results_state
        result_state = search_results['search_information']['organic_results_state']
        print(f"Organic result state: {result_state}",end="\n")

        if result_state == "Results for exact spelling":

            # Extracting links for positions 1, 2, or 3
            links = [result['link'] for result in search_results['organic_results'] if result['position'] in [1, 2, 3, 4 ,5]]

            # Looping through the links and printing them
            for index, url in enumerate(links):
                if check_domain_in_url(url, domain_name) or index < 5:
                    # Add data to list
                    Domain_name.append(domain_name)
                    Result_states.append(result_state)
                    URLs.append(url)
                    Status.append(True)


                    break
                elif index == 4:
                    break
                else:
                    continue

        else:
            # Add data to list
            Domain_name.append(domain_name)
            Result_states.append(result_state)
            URLs.append(None)
            Status.append(False)

        google_search_result = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Domain Name": Domain_name,
            "result_state": Result_states,
            "URL": URLs,
            "Status": Status
        }

        print(google_search_result)   
        save_result_sheet(google_search_result)


def get_excel_data(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name="input data")

        # Check if the column exists
        if "domain_name" not in df.columns:
            raise ValueError(f"Column 'domain_name' not found in the Excel file.")

        # Convert the column to a list
        data_list = df["domain_name"].dropna().tolist()  # Drop NaN values
        data_list = [domain.lower() for domain in data_list] # Convert to lowercase

        print(data_list)
        return data_list
    
    except Exception as e:
        print(f"An error occurred: {e}")
        return []
    
#-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

# Sheet name must be "input data"

excel_file_path = "C:/Users/shubh/Desktop/test.xlsx"

domain_list = get_excel_data(excel_file_path)

#google_search_result = google_search(domain_list)

google_search_result = serp_google_search(domain_list)

#save_result_sheet(google_search_result, excel_file_path)

print("""
            -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-= Code END =-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
""")

