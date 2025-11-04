import os
import re
# from docx import Document
import json


FILES_PATH = './data/raw/'

file_paths = [FILES_PATH+f for f in os.listdir(FILES_PATH) if re.match(r'.*\.docx', f)]

# print(file_paths)

# doc = Document(file_paths[1])

# print(doc)

from unstructured.partition.docx import partition_docx

# doc = file_paths[2]

# elements = partition_docx(filename=doc)
# print(doc)
# docs = []
# for el in elements:
#     docs.append({
#         "type": el.category,   # title, narrative_text, list_item, table, etc.
#         "text": el.text.strip()
#     })
#     print(el.category)
#     print(el.text.strip())
#     print('-'*80)

# print(docs)

all_docs_lst = []
for i, doc in enumerate(file_paths):
    # print('-'*10)
    # print(doc)
    # print('-'*10)
    elements = partition_docx(filename=doc)
    doc_lst = []
    for el in elements:
        if el.category == 'Title':
            doc_lst.append(f"[Title] {el.text.strip()}.\n")
        elif el.category == 'ListItem':
            doc_lst.append(f"- {el.text.strip()}.\n")
        else:
            doc_lst.append(el.text.strip())
    # all_docs_lst.append(' '.join(doc_lst))
    all_docs_lst.append({
        'doc_id': i ,
        'file_path': doc ,
        'text': ' '.join(doc_lst)
    })
        
# for i in all_docs_lst:
#     print(i)
#     print('-'*80)
    
with open("data/processed/proc_docx.jsonl", "w", encoding="utf-8") as f:
    for doc in all_docs_lst:
        json.dump(doc, f, ensure_ascii=False)
        f.write("\n")