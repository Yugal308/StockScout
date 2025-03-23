from flask import Flask, render_template, request, send_file
import os
import logging
from script import process_excel

app = Flask(__name__, 
            static_folder='static',
            template_folder='Templates')

app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Create uploads directory if it doesn't exist
for folder in [app.config['UPLOAD_FOLDER'], 'Templates']:
    if not os.path.exists(folder):
        os.makedirs(folder)
        logger.info(f"Created directory: {folder}")

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"Template error: {str(e)}")
        return "Template not found", 404

@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return "No file uploaded", 400
    
    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400
    
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        try:
            output_file = process_excel(file_path)
            return render_template('download.html', filename=output_file)
        except Exception as e:
            return f"Error processing file: {str(e)}", 500
        finally:
            # Clean up files
            if os.path.exists(file_path):
                os.remove(file_path)

@app.route('/download/<filename>')
def download_file(filename):
    try:
        path = os.path.join(os.getcwd(), filename)
        response = send_file(path, as_attachment=True)
        
        @response.call_on_close
        def remove_file():
            try:
                os.remove(path)
            except Exception as e:
                print(f"Error removing file: {e}")
        
        return response
    except Exception as e:
        return f"Error downloading file: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)