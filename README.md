# PDF Policy Extractor 

## Overview

This project was created to automate the process of manually extracting key information from insurance policy PDF documents. By utilizing Large Language Models (LLMs), it streamlines the data collection process, making it faster and less likely subject to human error. The goal is to efficiently build a structured dataset of policy details for analysis or record-keeping.

## What It Does

The `policy_extractor.py` script performs the following core functions:

1.  **PDF Content Extraction:** Reads text content directly from insurance policy PDF documents, both local files and URLs.
2.  **Data Extraction via LLMs:** Sends the extracted PDF text to a Large Language Model to intelligently identify and extract specific details such as:
    *   Payer Name
    *   Policy Name
    *   Policy Number
    *   Policy Type
    *   Approval Date
    *   Effective Date
    *   Brief Summary
    *   Covered HCPCS Codes (including CPT codes)
    *   Non-Covered HCPCS Codes (including CPT codes)
    *   Modifiers Used
    *   Service Summary
    *   Reimbursement Detail
    *   Non-Reimbursement Detail
    *   Notes
3.  **Structured Output to Excel:** Converts the extracted, structured information into a row in an Excel spreadsheet (`policy_data.xlsx`). Appends new policy data as additional rows to the same Excel file, allowing for the creation of a cumulative database of policy information.
5.  **User Interaction:** Guides the user through the process, prompting for necessary inputs (LLM API key, PDF source) and offering to open the generated Excel file.

## How to Use

### 1. Prerequisites

*   **Python:** Ensure you have Python installed on your system 
*   **Required Libraries:** Install the necessary Python packages by running the following command in your terminal:

    ```bash
    pip install PyPDF2 requests pandas openpyxl
    ```

*   **LLM API Key:** You will need an API key for a Large Language Model service. The script is configured to use OpenAI's GPT-4o by default, so an OpenAI API Key is recommended. You can get one from the [OpenAI Platform](https://platform.openai.com/account/api-keys).

### 2. Running the Program

1.  Save the `policy_extractor.py` file to your local machine.
2.  Open your terminal or command prompt.
3.  Navigate to the directory where you saved `policy_extractor.py`.
4.  Run the script using the Python interpreter:

    ```bash
    python policy_extractor.py
    ```

### 3. Interacting with the Program

The script will guide you through the process with command-line prompts:

*   **LLM API Key Prompt:**

    ```
    Please enter your LLM API Key:
    ```
    Enter your API key here. 

*   **Policy PDF Input Prompt:**

    ```
    Please enter the LOCAL path or URL to the insurance policy PDF (or 'q' to quit):
    ```
    *   **For a local PDF file:** Provide the full file path
    *   **For an online PDF:** Paste the direct URL to the PDF document 
    *   To exit the program, type `q` and press Enter.

### 4. Output and Continuous Use

Upon successful extraction, the program will:

*   Print the extracted JSON data to your terminal.
*   Create or update an Excel file named `policy_data.xlsx` in the same directory where the script is run. Each new policy processed will add a new row to this file.
*   Prompt you to open the Excel file:

    ```
    Do you want to open policy_data.xlsx now? (y/n):
    ```
    Type `y` to open the file automatically, or `n` to continue without opening.

The program will then return to the policy PDF input prompt, allowing you to process multiple documents in a single session. This enables you to build your `policy_data.xlsx` file incrementally.
