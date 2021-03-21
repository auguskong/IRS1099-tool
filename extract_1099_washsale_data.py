#!/usr/bin/env python
# coding: utf-8

# In[2]:


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


# In[3]:


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


# In[4]:


TEXT_ELEMENTS = [
    pdfminer.layout.LTTextBox,
    pdfminer.layout.LTTextBoxHorizontal,
    pdfminer.layout.LTTextLine,
    pdfminer.layout.LTTextLineHorizontal
]

FILE_NAME = "test.pdf"
COL_MARGIN = 0.5
page_layouts = extract_page_layouts(FILE_NAME)
print(page_layouts)
print("Number of pages: %d" % len(page_layouts))


# In[5]:


def extract_single_page_text(current_page):
    text = []
    for elem in current_page:
        if isinstance(elem, pdfminer.layout.LTTextBoxHorizontal):
            text.append(elem)
    return text


# In[6]:


raw_text = extract_single_page_text(page_layouts[0])
raw_text


# In[7]:


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


# In[8]:


raw_characters = extract_characters(raw_text)
raw_characters


# In[9]:


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

sorted_rows = arrange_text(raw_characters)
sorted_rows


# In[10]:


col_margin = 0.5
def create_separators(sorted_rows, margin):
    """Creates bounding boxes to fill the space between columns"""
    separators = []
    for row in sorted_rows:
        for idx, c in enumerate(row[:-1]): 
            if (row[idx+1].bbox[0] - c.bbox[2]) > margin:
                bbox = (c.bbox[2], c.bbox[3], row[idx+1].bbox[0], row[idx+1].bbox[1])
                separator = pdfminer.layout.LTRect(linewidth=2, bbox=bbox)
                separators.append(separator)
    return separators

separators = create_separators(sorted_rows, col_margin)


# In[11]:


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

text = arrange_and_extract_text(raw_characters)
text


# In[12]:



def generate_current_page_dataframe(text):
    columns = ['Date sold or disposed', 'Quantity', 'Proceeds', 'Date acquired', 'Cost or other basis', 'Wash sale loss disallowed (W)', 'Code', 'Gain or loss(-)', 'Additional information']
    df = pd.DataFrame(columns=columns)
    for row in text:
        if len(row) == 5:
            row = row.copy()
            row.insert(0, "-")
            row.insert(1, "-")
            row.insert(3, "-")
            row.insert(8, "-")
            temp = pd.DataFrame([row], columns=columns)
            df = df.append(temp, ignore_index=True)
        elif len(row) == 9:
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


# In[13]:


df = generate_current_page_dataframe(text)


# In[14]:


df


# In[ ]:




