
# Library Imports
import os
import json
import io
import platform
import PyPDF2
import requests
import pandas as pd

# Function to send prompt to LLM and receive a response
def call_llm(prompt, api_key):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-4o", 
        "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_object"} 
    }
    try:
        response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, json=data, timeout=120)
        response.raise_for_status() 
        return response.json()["choices"][0]["message"]["content"]
    except requests.exceptions.RequestException as e:
        print(f"Error calling LLM API: {e}")
        return None

# Function to read a PDF and extract data using LLM
def extract_policy_info(policy_input, api_key):
    text = ""
    
    # Decide whether to read from a URL or a local file
    if policy_input.startswith(("http://", "https://")):
        try:
            print(f"Attempting to download PDF from URL: {policy_input}")
            # Download the PDF content
            response = requests.get(policy_input, stream=True)
            response.raise_for_status() 
            # Read PDF content
            file_content = io.BytesIO(response.content)
            reader = PyPDF2.PdfReader(file_content)
        except requests.exceptions.RequestException as e:
            print(f"Error downloading PDF from URL: {e}")
            return
        except Exception as e:
            print(f"Error reading PDF from URL content: {e}")
            return
    else:
        try:
            print(f"Attempting to read PDF from local file: {policy_input}")
            # Open and read the local PDF file
            with open(policy_input, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
        except Exception as e:
            print(f"Error reading local PDF file: {e}")
            return

    # Extract text from the PDF
    for page_num in range(len(reader.pages)):
        text += reader.pages[page_num].extract_text()

    # Prompt LLM with text and instructions
    prompt = f"""
Act as a policy advisor. Please analyze the following insurance policy document and extract key details using the exact format below. Ensure each section is filled in accurately based on the content of the document.

Required Output Format:
Payer Name:
Policy Name:
Policy Number:
Policy Type:
Approval Date:
Effective Date:
Brief Summary:
Covered HCPCS Codes: (Include CPT codes as HCPCS codes)
Non-Covered HCPCS Codes: (Include CPT codes as HCPCS codes)
Modifiers Used:
Service Summary:
Reimbursement Detail:
Non-Reimbursement Detail:
Notes:

EXAMPLE:
Payer Name: Blue Cross Blue Shield NC
Policy Name: Anesthesia Services
Policy Number: AN1234
Policy Type: Anesthesia
Approval Date: 11/10/23
Effective Date: 1/1/24
Brief Summary: Covers anesthesia services when medically necessary during surgical procedures. Non-covered services include experimental anesthetic techniques.
Covered HCPCS Codes: ["85", "123", "456"]
Non Covered HCPCS Codes: ["12", "789"]
Modifiers Used: ["AA", "AD", "QK", "QX", "QY", "QZ"]
Service Summary: General, regional, and monitored anesthesia care for covered surgical procedures. Excludes experimental procedures.
Reimbursement Detail: Payment is based on ASA units multiplied by the contracted rate. Specific documentation requirements must be met.
Non-Reimbursement Detail: Experimental anesthesia, anesthesia for cosmetic procedures not covered
Notes: All claims must include proper ASA and CPT coding. Prior authorization required for specific high-cost procedures.

Please provide the entire output as a JSON object, directly, without any surrounding markdown code block fences (e.g., ```json). The JSON should have the following keys: payer_name, policy_name, policy_number, policy_type, approval_date, effective_date, brief_summary, covered_hcpcs_codes, non_covered_hcpcs_codes, modifiers_used, service_summary, reimbursement_detail, non_reimbursement_detail, notes. Ensure 'covered_hcpcs_codes' and 'non_covered_hcpcs_codes' are lists of strings (including CPT codes), and 'modifiers_used' is also a list of strings.

Policy Document Content:
{text}
    """

    print("\nCalling LLM with the extracted PDF content...")
    llm_response = call_llm(prompt, api_key)

    if llm_response:
        print(f"Raw LLM Response before parsing: {llm_response[:500]}...") 
        try:
            json_start = llm_response.find("```json")
            json_end = llm_response.rfind("```")

            if json_start != -1 and json_end != -1 and json_end > json_start:
                llm_response = llm_response[json_start + len("```json"):json_end].strip()
                print("Successfully stripped markdown code block fences from LLM response.")
            elif json_start != -1 and json_end == -1:
                llm_response = llm_response[json_start + len("```json"):].strip()
                print("Stripped only starting markdown code block fence from LLM response.")
            elif json_start == -1 and json_end != -1:
                llm_response = llm_response[:json_end].strip()
                print("Stripped only ending markdown code block fence from LLM response.")
            else:
                print("No markdown code block fences found or malformed fences. Attempting to parse as is.")

            # Parse the LLM response (which should now be pure JSON) into a dictionary
            extracted_data = json.loads(llm_response)
            print("\n--- Extracted Information (JSON) ---")
            print(json.dumps(extracted_data, indent=2))
            print("-------------------------------------")
            return extracted_data
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON from LLM response: {e}")
            print("LLM Response:")
            print(llm_response)
    return None

# Function to update the Excel sheet with extracted policy data
def update_excel_sheet(new_data, excel_file="policy_data.xlsx"):
    print("\n--- Attempting to update Excel sheet ---")
   
    data_to_excel = new_data.copy()
    # Convert list fields into comma-separated strings for Excel cells
    if 'covered_hcpcs_codes' in data_to_excel and isinstance(data_to_excel['covered_hcpcs_codes'], list):
        data_to_excel['covered_hcpcs_codes'] = ', '.join(data_to_excel['covered_hcpcs_codes'])
    if 'non_covered_hcpcs_codes' in data_to_excel and isinstance(data_to_excel['non_covered_hcpcs_codes'], list):
        data_to_excel['non_covered_hcpcs_codes'] = ', '.join(data_to_excel['non_covered_hcpcs_codes'])
    if 'modifiers_used' in data_to_excel and isinstance(data_to_excel['modifiers_used'], list):
        data_to_excel['modifiers_used'] = ', '.join(data_to_excel['modifiers_used'])

    df_new = pd.DataFrame([data_to_excel])
    print("DataFrame created from new data.")

    # Check if the Excel file already exists
    if os.path.exists(excel_file):
        print(f"Existing Excel file found: {excel_file}")
        try:
            # Read existing data and combine with the new row
            df_existing = pd.read_excel(excel_file)
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            print("New data concatenated with existing data.")
        except Exception as e:
            print(f"Error reading existing Excel file or concatenating: {e}")
            df_combined = df_new 
    else:
        print("No existing Excel file found. Creating a new one.")
        df_combined = df_new 
    try:
        df_combined.to_excel(excel_file, index=False) 
        print(f"Successfully updated {excel_file} with new policy data.")
        
        # Excel Styling
        try:
            from openpyxl import load_workbook
            from openpyxl.styles import Font, Alignment

            # Load the workbook and select the active sheet
            workbook = load_workbook(excel_file)
            sheet = workbook.active

            # Define columns that should have text wrapped
            wrap_text_columns = [
                "brief_summary", "service_summary", 
                "reimbursement_detail", "non_reimbursement_detail", "notes"
            ]

            # Map column headers to their letter 
            header_map = {cell.value.lower().replace(' ', '_'): cell.column_letter for cell in sheet[1]}

            # Set column widths for better readability and apply text wrapping
            for col in sheet.columns:
                max_length = 0
                column_letter = col[0].column_letter 
                column_header = sheet[column_letter + '1'].value.lower().replace(' ', '_') 

                for cell in col:
                    try: 
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                        # Apply text wrapping for identified columns
                        if column_header in wrap_text_columns:
                            cell.alignment = Alignment(wrap_text=True)
                    except:
                        pass
                
                adjusted_width = (max_length + 2) # Add a little padding
                
                #Cap the width to a reasonable maximum for text-wrapped columns
                if column_header in wrap_text_columns:
                    adjusted_width = min(adjusted_width, 70) 

                sheet.column_dimensions[column_letter].width = adjusted_width

            # Make the header row bold
            for cell in sheet['1:1']:
                cell.font = Font(bold=True)

            # Save the changes to the Excel file
            workbook.save(excel_file)
            print("Applied basic styling and text wrapping to the Excel sheet.")

        except Exception as e:
            print(f"Error applying Excel styling: {e}")

        # Ask the user if they want to open the Excel file
        open_excel = input(f"Do you want to open {excel_file} now? (y/n): ").lower()
        if open_excel == 'y':
            try:
                # Open the Excel file based on the operating system
                if platform.system() == "Windows":
                    os.startfile(excel_file)
                elif platform.system() == "Darwin": # macOS
                    os.system(f"open {excel_file}")
                else: # Linux 
                    os.system(f"xdg-open {excel_file}")
            except Exception as e:
                print(f"Error opening Excel file: {e}")

    except Exception as e:
        print(f"Error writing to Excel file: {e}")

# Main function to run the program
def main():
    """This is the main function that runs the policy extraction program.
    It handles user input for the API key and PDF documents, and manages the extraction process.
    """
    # Get LLM API key. Prompt user if not found in environment variables
    api_key = os.getenv("LLM_API_KEY")
    if not api_key:
        api_key = input("Please enter your LLM API Key: ")
    
    # Loop to allow processing multiple policy documents continuously
    while True: 
        # Ask the user for the PDF path or URL
        policy_path = input("Please enter the LOCAL path or URL to the insurance policy PDF (or 'q' to quit): ")
        if policy_path.lower() == 'q':
            break 

        # Validate if the input is a valid PDF file or URL
        if not policy_path.startswith(("http://", "https://")):
            # Check if local file exists
            if not os.path.exists(policy_path):
                print(f"Error: File not found at {policy_path}")
                continue 
            
            # Check if local file is a PDF
            if not policy_path.lower().endswith('.pdf'):
                print("Error: The provided file does not seem to be a PDF. Please provide a PDF file or URL.")
                continue 
        else:
            # Check if URL points to a PDF
            if not policy_path.lower().endswith('.pdf'):
                print("Error: The provided URL does not seem to point to a PDF. Please provide a PDF URL.")
                continue 

        # Extract information and update the Excel sheet if successful
        extracted_info = extract_policy_info(policy_path, api_key)
        if extracted_info:
            update_excel_sheet(extracted_info)
        print("\n") 

# Run the main function when the script is executed
if __name__ == "__main__":
    main()