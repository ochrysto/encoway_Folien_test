from pprint import pprint
from uuid import uuid4

import ollama
from llama_index.core import SimpleDirectoryReader


print("Create a httpx ollama client")
client = ollama.Client(host='http://localhost:11434')
source = 'C:\\Users\chrystosseko\Documents\Abschlussarbeit\encoway_Folien_test\Source'

# liest pdf, docs, csv, jpg, md, pptx
# read document
print("Instantiate SimpleDirectoryReader")

reader = SimpleDirectoryReader(source, recursive=True)
print("Load data")
docs = reader.load_data(show_progress=True)
pprint(docs)

# Write an output into a file
with open(f"output_{uuid4()}.txt", mode='w') as output_file:
    first_doc = docs[0]
    output_file.write(first_doc.text)

