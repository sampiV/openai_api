# Excel Sheet Summarizer using OpenAI GPT

This Python script reads all sheets from an Excel file and uses OpenAI's GPT API to generate concise text summaries of each sheet's contents. It's especially useful for quickly understanding large or multi-sheet datasets.

---

## Features

- Summarizes each sheet in an Excel file using a language model
- Supports `.xlsx` files with multiple sheets
- Uses OpenAI's GPT-3.5-turbo-instruct (or similar) for text generation
- Automatically handles API and data reading errors gracefully

---

## How It Works

1. Loads your Excel file using `pandas`
2. Converts each sheet to a text format
3. Sends a prompt to the OpenAI API to generate a summary
4. Prints the summary of each sheet in the terminal

---

## Requirements

- Python 3.7+
- `openai`
- `pandas`

You can install the required packages with:

```bash
pip install openai pandas
