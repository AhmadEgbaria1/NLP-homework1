# 🏛️ NLP Homework 1: Building a Corpus from Knesset Protocols

![Data Mining](https://img.shields.io/badge/Task-Corpus_Building-orange)
![Python](https://img.shields.io/badge/Python-3.8+-green)
![Language](https://img.shields.io/badge/Language-Hebrew-blue)

This repository contains the implementation for Homework 1, focusing on the creation of a linguistic corpus from official Israeli Knesset session protocols. The project involves large-scale data extraction from unstructured document formats.

---

## 📋 Project Overview

The core of this assignment was to transform hundreds of raw `.docx` files containing parliamentary discussions into a structured, searchable, and analyzable text corpus. 

### Key Challenges & Features:
* **DOCX Parsing:** Implementing automated scripts to batch-process large numbers of Microsoft Word documents and extract raw text content.
* **Hebrew Text Processing:** Handling the unique challenges of the Hebrew language, including Right-to-Left (RTL) formatting and specific morphological structures.
* **Text Cleaning:** Using Regular Expressions (Regex) to strip away metadata, timestamps, and non-linguistic noise from the protocols.
* **Corpus Statistics:** Building a frequency-based vocabulary and calculating distribution metrics for the words used in the Knesset sessions.

---

## 📂 Repository Structure

* `docx_parser.py` - Logic for iterating through folders and extracting text from `.docx` files.
* `cleaner.py` - Regex-based cleaning pipeline specifically tuned for Hebrew parliamentary text.
* `corpus_builder.py` - Functions to tokenize the text and generate the final word corpus and frequency dictionaries.
* `main.py` - The main execution script to process the entire dataset.

---

## 🚀 How to Run

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/AhmadEgbaria1/NLP-homework1.git](https://github.com/AhmadEgbaria1/NLP-homework1.git)
   cd NLP-homework1
