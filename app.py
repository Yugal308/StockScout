from flask import Flask, render_template, request, send_file
import os
from script import process_excel

app = Flask(__name__, static_folder='static')
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# Create uploads directory if it doesn't exist
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return render_template('index.html')

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