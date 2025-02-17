# Flask App: Upload Excel and Generate PowerPoint
# Folder Structure:
# project_root/
#   ┣ templates/
#   ┃   ┗ upload.html  (HTML for upload form)
#   ┣ uploads/         (Created automatically for user uploads)
#   ┣ outputs/         (Created automatically for generated files)
#   ┣ app.py           (Flask routes and endpoints)
#   ┣ processing/
#   ┃   ┣ __init__.py   (Makes it a package)
#   ┃   ┗ excel_to_ppt.py (Functions to process Excel to PPT)
#   ┗ requirements.txt  (Dependencies)


from flask import Flask, render_template, request, send_file
from excel_to_ppt.processor import generate_ppt
import os
import zipfile


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'data/input'

@app.route('/')
def upload_form():
    return render_template('upload.html')

@app.route('/generate-ppt', methods=['POST'])
def generate_ppt_endpoint():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    if file:
        input_file = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(input_file)
        output_path = 'data/output'
        template_path = 'data/input/template.pptx'
        image_path = 'data/input/images'
        output_files = generate_ppt(input_file, output_path, template_path, image_path)
        
        # Create a zip file containing all output files
        zip_path = os.path.join(output_path, f'ACB_{file.filename}.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in output_files:
                zipf.write(file, os.path.basename(file))
        
        return send_file(zip_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
