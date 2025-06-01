# app.py
from flask import Flask, request, render_template, send_file, redirect, url_for
import os
import pandas as pd  # For Excel handling example
from LVCalculator import run_processor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSED_FOLDER'] = 'processed'
app.config['LV_FOLDER'] = 'LV'

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if file was uploaded
        if 'file' not in request.files:
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '' or file.filename[-4:] != '.pdf':
            return redirect(request.url)
        
        if file:
            # Save uploaded file
            upload_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(upload_path)
            
            # Process file (replace with your Python script)
            output_path = run_processor(app.config['LV_FOLDER'], 
                                        upload_path, 
                                        app.config['PROCESSED_FOLDER'],
                                        file.filename)
            
            # Redirect to download page
            return redirect(url_for('download_file', filename=os.path.basename(output_path)))
    
    return render_template('upload.html')

@app.route('/download/<filename>')

def download_file(filename):
    return send_file(
        os.path.join(app.config['PROCESSED_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )

# def process_file(input_path):
#     # Replace this with your actual Python processing script
#     # Example: Read CSV, process, save as Excel
#     df = pd.read_csv(input_path)
#     processed_df = df * 2  # Example processing
    
#     output_path = os.path.join(app.config['PROCESSED_FOLDER'], 'processed.xlsx')
#     processed_df.to_excel(output_path, index=False)
    
#     return output_path

if __name__ == '__main__':
    app.run(debug=True)