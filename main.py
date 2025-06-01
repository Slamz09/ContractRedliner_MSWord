# app.py
from flask import Flask, request, render_template_string, send_file
import os
import time
from review_contract import review_and_edit_contract

app = Flask(__name__)
UPLOAD_FOLDER = r"C:\Users\slamb\PythonProjects\contract_review_ai"  # Your project folder
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
