import openai
import pandas as pd

# Set your OpenAI API key (replace with your actual key or set as an environment variable)
openai.api_key = 'your-openai-api-key-here'

def summarize_sheet(sheet_data, sheet_name):
    """
    Summarize the data from a given sheet using the OpenAI API.
    """
    sheet_content = sheet_data.to_string(index=False)
    prompt = f"Summarize the following data from the Excel sheet '{sheet_name}':\n\n{sheet_content}\n\nSummary:"

    try:
        response = openai.Completion.create(
            model="gpt-3.5-turbo-instruct",  # Note: Updated to newer instruct-compatible model
            prompt=prompt,
            max_tokens=100,
            temperature=0.7
        )
        summary = response.choices[0].text.strip()
        return summary
    except Exception as e:
        return f"Error during API call: {e}"

def summarize_excel(file_path):
    """
    Summarize each sheet in the Excel file.
    """
    try:
        excel_data = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return

    for sheet_name in excel_data.sheet_names:
        try:
            sheet_data = pd.read_excel(excel_data, sheet_name=sheet_name)
            summary = summarize_sheet(sheet_data, sheet_name)
            print(f"\nSummary for sheet '{sheet_name}':\n{summary}\n")
        except Exception as e:
            print(f"Error processing sheet '{sheet_name}': {e}")

if __name__ == "__main__":
    file_path = input("Please enter the path to the Excel file: ")
    summarize_excel(file_path)
