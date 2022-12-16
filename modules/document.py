from docx import Document
from pathlib import Path
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# https://python-docx.readthedocs.io/en/latest/user/hdrftr.html

# python concurrent.future

file_name = "demo.docx"
file_path = Path(__file__).parent.parent / file_name

def main():

    # for i in range(len(documents)):
    #     document = add_footer(documents[i], f"this is the {i}th document")
    #     document.save(f"{i}th_file.docx")

    documents = []
    future = []
    with ThreadPoolExecutor() as executor:
        for _ in range(200):
            future.append(executor.submit(Document))

        for future in as_completed(future):
            documents.append(future.result())

    print("here")

    future = []
    num = 0
    with ThreadPoolExecutor() as executor:
        for i in range(len(documents)):
            future.append(executor.submit(add_footer, documents[i], f"this is the {i}th document"))
        for future in as_completed(future):
            future.result().save(f"{num}th_file.docx")
            num += 1

    print(time.perf_counter())

def sleep_for_a_bit(second):
    print(f"sleeping {second} seconds")
    time.sleep(second)
    print(f"Finished sleeping {second} seconds")
    return "Done sleeping"

def add_footer(document, footer_text):
    section = document.sections[0]
    footer = section.footer
    footer_paragraph = footer.paragraphs[0]
    footer_paragraph.text = footer_text
    return document

def open_file(file_path):
    os.system("start " + str(file_path))

if __name__ == "__main__":
    main()