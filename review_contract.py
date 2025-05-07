# review_contract.py
from docx import Document
import spacy
import win32com.client as win32
import pythoncom
import os
import re
from playbook import playbook

# Load spaCy model
nlp = spacy.load("en_core_web_sm")

def determine_contract_type(doc):
    full_text = " ".join(para.text.strip().lower() for para in doc.paragraphs if para.text.strip())
    customer_score = sum(1 for kw in playbook["payment_customer"]["context_keywords"] if kw in full_text)
    vendor_score = sum(1 for kw in playbook["payment_vendor"]["context_keywords"] if kw in full_text)
    print(f"Customer score: {customer_score}, Vendor score: {vendor_score}")
    return "customer" if customer_score > vendor_score else "vendor" if vendor_score > customer_score else None

def review_and_edit_contract(file_path):
    file_path = os.path.abspath(file_path)
    print(f"Processing file: {file_path}")
    doc = Document(file_path)
    contract_type = determine_contract_type(doc)
    print(f"Determined contract type: {contract_type or 'Unknown'}")

    pythoncom.CoInitialize()
    word = None
    word_doc = None
    try:
        word = win32.Dispatch("Word.Application")
        # Do not set Visible property; let it run in background
        word_doc = word.Documents.Open(file_path)
        word_doc.TrackRevisions = True
        print("TrackRevisions enabled: ", word_doc.TrackRevisions)

        for i, para in enumerate(word_doc.Paragraphs, 1):
            text = para.Range.Text.strip()
            if not text:
                continue
            print(f"\nProcessing paragraph {i}: '{repr(text)}'")
            text_lower = text.lower()

            edited = False
            for clause_type, rules in playbook.items():
                print(f"  Checking for '{clause_type}' with keywords: {rules['keywords']}")
                matched_keyword = next((kw for kw in rules["keywords"] if kw in text_lower), None)
                if matched_keyword:
                    print(f"  Match found for '{clause_type}' with keyword '{matched_keyword}'")

                    if clause_type.startswith("payment"):
                        if contract_type == "customer" and clause_type != "payment_customer":
                            continue
                        elif contract_type == "vendor" and clause_type != "payment_vendor":
                            continue
                        elif not contract_type:
                            if clause_type != "payment_customer":
                                continue
                            print("  Warning: Unknown contract type, defaulting to customer payment terms.")

                        preferred_days = "30" if clause_type == "payment_customer" else "60"
                        match = re.search(r"(\d+)\s*days", text_lower)
                        if match and match.group(1) != preferred_days:
                            old_days = match.group(1)  # e.g., "90"
                            start_pos = match.start(1)
                            end_pos = match.end(1)
                            range_to_edit = para.Range.Duplicate
                            range_to_edit.SetRange(para.Range.Start + start_pos, para.Range.Start + end_pos)
                            print(f"  Replacing '{old_days}' with '{preferred_days}' at range {start_pos}-{end_pos}")
                            range_to_edit.Text = preferred_days
                            edited = True

                    elif clause_type == "termination":
                        current_party = "Customer"
                        preferred_party = "Either party"
                        if current_party.lower() in text_lower and preferred_party.lower() not in text_lower:
                            start_pos = text_lower.find(current_party.lower())
                            end_pos = start_pos + len(current_party)
                            range_to_edit = para.Range.Duplicate
                            range_to_edit.SetRange(para.Range.Start + start_pos, para.Range.Start + end_pos)
                            print(f"  Replacing '{current_party}' with '{preferred_party}' at range {start_pos}-{end_pos}")
                            range_to_edit.Text = preferred_party
                            edited = True

            if edited:
                print(f"  After edit, paragraph text: '{para.Range.Text.strip()}'")
            else:
                print("  No changes made to this paragraph.")

        # Verify changes before saving
        print("\nVerifying document content before saving:")
        for i, para in enumerate(word_doc.Paragraphs, 1):
            print(f"Paragraph {i}: '{repr(para.Range.Text.strip())}'")

        output_path = file_path.replace(".docx", "_legal_redline.docx")
        word_doc.SaveAs(output_path)
        word_doc.Close()
        word.Quit()
        print(f"Saved successfully to: '{output_path}'")
        return output_path

    except Exception as e:
        print(f"Error during processing: {str(e)}")
        if 'word_doc' in locals() and word_doc:
            word_doc.Close()
        if 'word' in locals() and word:
            word.Quit()
        return None

    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    file_path = r"C:\Users\slamb\PythonProjects\contract_review_ai\sample_contract.docx"
    reviewed_file = review_and_edit_contract(file_path)
    if reviewed_file:
        print(f"Reviewed contract saved as: {reviewed_file}")
    else:
        print("Failed to save reviewed contract.")