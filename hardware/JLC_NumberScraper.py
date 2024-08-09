#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Created on %(date)s

@author: Chance Fleeting

This code is formatted in accordance with PEP8.
See: https://peps.python.org/pep-0008/

Use %matplotlib qt to display images in a pop-out window.
Use %matplotlib inline to display images inline.
"""

from __future__ import annotations

__author__ = 'Chance Fleeting'
__version__ = '0.2'

import pandas as pd
import re
import requests
from bs4 import BeautifulSoup
from functools import lru_cache

# %% CONSTANTS
JLCPCB_URL = "https://jlcpcb.com/parts/componentSearch?searchTxt="
HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
}
MAXMEMO = 100

# %% FUNCTIONS
@lru_cache(maxsize=100)
def scrape_jlcpcb_component_code(product_name: str) -> str | None:
    """
    Scrape the JLCPCB website to find the component code for a given product name.

    Parameters
    ----------
    product_name : str
        The product name to search for on the JLCPCB website.

    Returns
    -------
    str or None
        The JLCPCB component code if found, otherwise None.
    """
    search_url = JLCPCB_URL + product_name.replace(" ", "%20")
    response = requests.get(search_url, headers=HEADERS)
    soup = BeautifulSoup(response.text, 'html.parser')

    # HACK: Use regex to extract the component code from the HTML
    match_val = re.search(r'componentCode:"([^"]+)', soup.__repr__())
    
    if match_val:
        return match_val.group(1)
    else:
        print(f"Could not find component code for product: {product_name}")
        return None
    
@lru_cache(maxsize=100)
def process_spreadsheet(input_file: str, output_file: str) -> None:
    """
    Process an input spreadsheet containing product names, scrape JLCPCB for
    component codes, and save the results to an output spreadsheet.

    Parameters
    ----------
    input_file : str
        Path to the input spreadsheet (Excel file) containing product names.
    output_file : str
        Path to save the output spreadsheet with component codes added.
    """
    df = pd.read_excel(input_file)

    if 'Manufacturer Part #' not in df.columns:
        raise ValueError("Input spreadsheet must contain a 'Manufacturer Part #' column.")

    df['ComponentCode'] = df['Manufacturer Part #'].apply(scrape_jlcpcb_component_code)

    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

# %% MAIN EXECUTION
if __name__ == "__main__":
    """
    Main execution block for the script.
    """
    input_file = "BOM.xlsx"  # Replace with your input file path
    output_file = "JLC_NumberScraper_output.xlsx"  # Replace with your desired output file path

    #start_time = time.time()
    process_spreadsheet(input_file, output_file)
    #elapsed_time = time.time() - start_time

    #print(f"Script completed in {elapsed_time:.2f} seconds.")
