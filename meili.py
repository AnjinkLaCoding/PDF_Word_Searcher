#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import fitz
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


# In[2]:


def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = "".join(page.get_text("text") for page in doc)
    return text


# In[3]:


def index_pdfs():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, folder_selected)
    folder_path = entry_pdf.get()
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            text = extract_text_from_pdf(pdf_path)
            pdf_data.append({"id": filename, "filename": filename, "content": text})
    
    if pdf_data:
        messagebox.showinfo("Success", "PDFs indexed successfully!")
    else:
        messagebox.showwarning("Warning", "No PDFs found in the selected folder.")


# In[4]:


def search_word():
    found_files = []
    query = entry_words.get()
    if not query:
        messagebox.showinfo("Results", "Please enter a word")
        return

    for pdf in pdf_data:
        if query.lower() in pdf["content"].lower():
            found_files+=[pdf["filename"]]
    
    if not found_files:
        messagebox.showinfo("Results", "The word can't be found!")
        return

    messagebox.showinfo("Results", "Word found in:\n" + "\n".join(found_files))


# In[5]:


def save_as():
    found_files = []
    query = entry_words.get()
    if not query:
        messagebox.showinfo("Results", "Please enter a word")
        return

    for pdf in pdf_data:
        if query.lower() in pdf["content"].lower():
            found_files+=[pdf["filename"]]
    
    if not found_files:
        messagebox.showinfo("Results", "The word can't be found!")
        return

    #messagebox.showinfo("Results", "Word found in:\n" + "\n".join(found_files))
    save_path = entry_output.get()
    if save_path:
        with open(save_path, "w") as file:
            file.write("Search Results for the word " + query + ":\n\n")
            file.write("\n".join(found_files))
        messagebox.showinfo("Saved", "Results saved successfully!")
    else:
        messagebox.showwarning("Warning", "Error saving file!")


# In[6]:


def select_output():
    file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, file_path)


# In[7]:


# GUI Setup
root = tk.Tk()
root.title("PDF Word Searcher")
root.geometry("400x250")

pdf_data = []

tk.Label(root, text="Select Folder:").pack()
entry_pdf = tk.Entry(root, width=50) ##variable to pass the user input to the function
entry_pdf.pack()
tk.Button(root, text="Browse", command=index_pdfs).pack()

tk.Label(root, text="Enter words to search:").pack()
entry_words = tk.Entry(root, width=50)
entry_words.pack()

tk.Label(root, text="Save As:").pack()
entry_output = tk.Entry(root, width=50)
entry_output.pack()
tk.Button(root, text="Browse", command=select_output).pack()

tk.Button(root, text="Generate List file", command=save_as).pack()

root.mainloop()

