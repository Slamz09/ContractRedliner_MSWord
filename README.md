
******Contract Redliner Set Up Guide******


**1. Create the main project folder and all project files shown below:**

         contract_rediner
         ├── app.py
         ├── requirements.txt
         ├── review_contract.py
         ├── playbook.py
         └── sample_contract.docx (optional test file)

***Paste each of the codes below into Notepad or Notepad++ to create the above listed project files***
    
**app.py**           
     
      from flask import Flask, request, render_template_string, send_file
      import os
      import time
      from review_contract import review_and_edit_contract

      app = Flask(__name__)
      UPLOAD_FOLDER = r"C:\Users\slamb\PythonProjects\contract_redliner"  # Your project folder
      app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

      # HTML template as a string
      HTML_TEMPLATE = """
      <!DOCTYPE html>
      <html>
      <head>
    <title>Upload Contract for Review</title>
      </head>
      <body>
    <h1>Upload Contract for Review</h1>
    <form method="post" enctype="multipart/form-data">
        <input type="file" name="contract" accept=".docx">
        <input type="submit" value="Review">
          </form>
       {% if message %}
           <p>{{ message }}</p>
       {% endif %}
       {% if redline_file %}
           <p>Download your redlined contract: <a href="/download/{{ redline_file }}">{{ redline_file }}</a></p>
       {% endif %}
         </body>
         </html>
         """
      @app.route("/", methods=["GET", "POST"])
      def upload_contract():
    message = None
    redline_file = None
    
    if request.method == "POST":
        # Check if a file was uploaded
        if "contract" not in request.files:
            message = "No file uploaded."
            return render_template_string(HTML_TEMPLATE, message=message)
        
        file = request.files["contract"]
        if file.filename == "":
            message = "No file selected."
            return render_template_string(HTML_TEMPLATE, message=message)
        
        if file and file.filename.endswith(".docx"):
            # Generate a unique filename with timestamp
            timestamp = int(time.time())
            original_filename = f"contract_{timestamp}_{file.filename}"
            original_path = os.path.join(app.config["UPLOAD_FOLDER"], original_filename)
            
            # Save the uploaded file
            file.save(original_path)
            print(f"Saved uploaded file as: {original_path}")
            
            # Process with AI and get redlined path
            redline_path = review_and_edit_contract(original_path)
            if redline_path:
                redline_file = os.path.basename(redline_path)  # Just the filename for download
                message = f"Contract processed successfully. Original saved as: {original_filename}"
            else:
                message = "Error processing the contract."
    
    return render_template_string(HTML_TEMPLATE, message=message, redline_file=redline_file) 
    @app.route("/download/<filename>")
    def download_file(filename):
    # Serve the redlined file for download
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    return send_file(file_path, as_attachment=True)
    if __name__ == "__main__":
    app.run(debug=True)  
    
  
**requirements.txt**

      flask
      python-docx
      pywin32
      spacy


**review_contract.py**

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


**playbook.py**

      playbook = {
    "payment_customer": {
        "keywords": ["payment", "billing", "invoice", "compensation", "terms of payment"],
        "context_keywords": ["customer shall pay", "client shall pay", "net 30"],
        "preferred": "Payment terms shall be Net 30 days.",
    },
    "payment_vendor": {
        "keywords": ["payment", "billing", "invoice", "compensation", "terms of payment"],
        "context_keywords": ["vendor shall be paid", "clean harbors shall pay", "net 60"],
        "preferred": "Payment terms shall be Net 60 days.",
    },
    "termination": {
        "keywords": ["termination", "cancel", "end", "notice", "termination rights"],
        "preferred": "Either party may terminate with 30 days' written notice.",
    }
}

**2. Open power shell, type 'cd' and then paste the file path of the contract_redliner project folder.**
   
            cd C:\Users\slamb\PythonProjects\contract_redliner
   
**4. Run the app.py file:**
   
            python app.py
   
**6. Open browser and go to http://127.0.0.1:5000/.**
   
**8. Upload contract and click ‘review’**
   
**10. Open the newly created redlined contract doc in contract_redliner project folder.**




        
