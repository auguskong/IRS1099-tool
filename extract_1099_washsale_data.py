#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pdfquery import PDFQuery
import pdfminer
from pdfminer.pdfpage import PDFPage, PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LTRect, LTLine

import matplotlib.pyplot as plt
from matplotlib import patches
get_ipython().run_line_magic('matplotlib', 'inline')

import pandas as pd


# In[45]:


def extract_page_layouts(file):
    """
    Extracts LTPage objects from a pdf file.
    modified from: http://www.degeneratestate.org/posts/2016/Jun/15/extracting-tabular-data-from-pdfs/
    Tests show that using PDFQuery to extract the document is ~ 5 times faster than pdfminer.
    """
    laparams = LAParams()
    
    with open(file, mode='rb') as pdf_file:
        print("Open document %s" % pdf_file.name)
        document = PDFQuery(pdf_file).doc

        if not document.is_extractable:
            raise PDFTextExtractionNotAllowed

        rsrcmgr = PDFResourceManager()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        layouts = []
        for page in PDFPage.create_pages(document):
            interpreter.process_page(page)
            layouts.append(device.get_result())
    
    return layouts


# In[46]:


TEXT_ELEMENTS = [
    pdfminer.layout.LTTextBox,
    pdfminer.layout.LTTextBoxHorizontal,
    pdfminer.layout.LTTextLine,
    pdfminer.layout.LTTextLineHorizontal
]


# In[47]:


def extract_single_page_text(current_page):
    text = []
    for elem in current_page:
        if isinstance(elem, pdfminer.layout.LTTextBoxHorizontal):
            text.append(elem)
    return text


# In[48]:


def flatten(lst):
    """Flattens a list of lists"""
    return [item for sublist in lst for item in sublist]

def extract_characters(element):
    """
    Recursively extracts individual characters from 
    text elements. 
    """
    if isinstance(element, pdfminer.layout.LTChar):
        return [element]

    if any(isinstance(element, i) for i in TEXT_ELEMENTS):
        return flatten([extract_characters(e) for e in element])

    if isinstance(element, list):
        return flatten([extract_characters(l) for l in element])

    return []


# In[49]:


def arrange_text(characters):
    """
    For each row find the characters in the row
    and sort them horizontally.
    """
    
    # find unique y0 (rows) for character assignment
    rows = sorted(list(set(c.bbox[1] for c in characters)), reverse=True)
    
    sorted_rows = []
    for row in rows:
        sorted_row = sorted([c for c in characters if c.bbox[1] == row], key=lambda c: c.bbox[0])
        sorted_rows.append(sorted_row)
    return sorted_rows


# In[50]:


def arrange_and_extract_text(characters, margin=0.5):
    
    rows = sorted(list(set(c.bbox[1] for c in characters)), reverse=True)
    
    row_texts = []
    for row in rows:
        
        sorted_row = sorted([c for c in characters if c.bbox[1] == row], key=lambda c: c.bbox[0])
        
        col_idx=0
        row_text = []
        for idx, char in enumerate(sorted_row[:-1]):
            if (sorted_row[idx+1].bbox[0] - char.bbox[2]) > margin:
                col_text = "".join([c.get_text() for c in sorted_row[col_idx:idx+1]])
                col_idx = idx+1
                row_text.append(col_text)
            elif idx==len(sorted_row)-2:
                col_text = "".join([c.get_text() for c in sorted_row[col_idx:]])
                row_text.append(col_text) 
        row_texts.append(row_text)
    return row_texts


# In[51]:


def generate_current_page_dataframe(text):
    columns = ['Date sold or disposed', 'Quantity', 'Proceeds', 'Date acquired', 'Cost or other basis', 'Wash sale loss disallowed (W)', 'Code', 'Gain or loss(-)', 'Additional information']
    df = pd.DataFrame(columns=columns)
    for row in text:
        if len(row) == 5 and row[3] == "W":
            row = row.copy()
            row.insert(0, "-")
            row.insert(1, "Total")
            row.insert(3, "-")
            row.insert(8, "-")
            temp = pd.DataFrame([row], columns=columns)
            df = df.append(temp, ignore_index=True)
        elif len(row) == 9 and row[6] == "W":
            df_row = pd.DataFrame([row], columns=columns)
            df = df.append(df_row, ignore_index=True)
        elif len(row) == 8 and row[5] == "W":
            row = row.copy()
            row.insert(3, "Various")
            temp = pd.DataFrame([row], columns=columns)
            df = df.append(temp, ignore_index=True)
        else:
            continue
    return df


# In[52]:


def extract_data_from_one_page(page):
    raw_text = extract_single_page_text(page)
    # print("raw_text", raw_text)
    raw_characters = extract_characters(raw_text)
    # print("raw_characters", raw_characters)
    sorted_rows = arrange_text(raw_characters)
    # print("sorted_rows", sorted_rows)
    text = arrange_and_extract_text(raw_characters)
    df = generate_current_page_dataframe(text)
    return df


# In[53]:


def extract_data_from_all_pages(pages):
    columns = ['Date sold or disposed', 'Quantity', 'Proceeds', 'Date acquired', 'Cost or other basis', 'Wash sale loss disallowed (W)', 'Code', 'Gain or loss(-)', 'Additional information']
    df_total = pd.DataFrame(columns=columns)
    for p in pages:
        df_curr = extract_data_from_one_page(p)
        df_total = df_total.append(df_curr)
    return df_total


# In[90]:


FILE_NAME = "1099-TD.pdf"
COL_MARGIN = 0.5
page_layouts = extract_page_layouts(FILE_NAME)
df_total_raw = extract_data_from_all_pages(page_layouts)
df_clean = df_total_raw.loc[df_total_raw['Quantity'] != 'Total']


# In[92]:


df_match_8949_form = df_clean[['Quantity', 'Date acquired', 'Date sold or disposed', 'Proceeds', 'Cost or other basis', 'Code', 'Wash sale loss disallowed (W)','Gain or loss(-)']]

wash_sale_data = df_match_8949_form.to_dict('records')


# In[97]:


df_clean_sorted = df_clean.reset_index(drop=True) #use the default index


# In[99]:


def total_value_for_each_form_page(col_name):
    df_col = df_clean_sorted[col_name]
    df_col.convert_objects(convert_numeric=True)

    dp_no_comma = df_col.apply(lambda row: row.replace(',', ''))
    dp_no_comma = dp_no_comma.convert_objects(convert_numeric=True)
    N = 14
    
    return dp_no_comma.groupby(dp_no_comma.index // N).sum(level=0)


# In[ ]:


total_value_for_each_form_page('Proceeds')


# In[ ]:


total_value_for_each_form_page('Cost or other basis')


# In[ ]:


total_value_for_each_form_page('Wash sale loss disallowed (W)')


# In[ ]:


total_value_for_each_form_page('Gain or loss(-)')


# In[ ]:


import pyautogui as pgui
import pyperclip as pclip
import time

x_coordinates_dict = {
    "date_acquired": 600,
    "date_sold": 650,
    "proceeds": 700,
    "cost_basis": 780,
    "code_from_instructions": 840,
    "amount_of_adjustment": 900,
    "gain_or_loss": 980
}

def fill_8949_form():
    first_row_y_coordinate = 420
    row_interval = 26
    num_rows = 14 # form 8949
    
    curr_data_index = 28
    curr_row_index = 0

    x_coordinates = [500, 600, 650, 700, 780, 840, 900, 980]
    
    pgui.click(x_coordinates[0], first_row_y_coordinate, interval=0.5)
    wash_sale_data_keys = list(wash_sale_data[0])

    print("Start filling the form")
    while curr_row_index < num_rows:
        curr_row_y_coordinate = first_row_y_coordinate + row_interval * curr_row_index
        print("filling form row", curr_row_index + 1, " data index is: ", curr_data_index)
        for col_idx, col_x_coordinate in enumerate(x_coordinates):
            pgui.click(col_x_coordinate, curr_row_y_coordinate, interval=0.2)
            time.sleep(1)
            pclip.copy(wash_sale_data[curr_data_index][wash_sale_data_keys[col_idx]])
            pgui.hotkey('command', 'v')
        curr_row_index += 1
        curr_data_index += 1
    print("complete filling the form")
       

