from flask import Flask, request, jsonify, send_file
import os
from generate_ppt import generate_ppt

app = Flask(__name__)

# Create a directory for storing presentations
UPLOAD_FOLDER = 'generated_presentations'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return '''
    <html>
        <head>
            <title>Decklyst - Generate Presentations</title>
            <style>
                body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
                .form-group { margin-bottom: 15px; }
                label { display: block; margin-bottom: 5px; }
                input { width: 100%; padding: 8px; margin-bottom: 10px; }
                button { background: #4CAF50; color: white; padding: 10px 15px; border: none; cursor: pointer; }
                button:hover { background: #45a049; }
                button:disabled { background: #cccccc; cursor: not-allowed; }
                #result { margin-top: 20px; padding: 10px; }
                .error { color: #ff0000; }
                .success { color: #4CAF50; }
                .download-link { display: inline-block; margin-top: 10px; color: #0066cc; text-decoration: none; }
                .download-link:hover { text-decoration: underline; }
            </style>
        </head>
        <body>
            <h1>Decklyst - Generate Presentations</h1>
            <div class="form-group">
                <label for="prompt">Prompt (write me a PPT presentation about...):</label>
                <input type="text" id="prompt" name="prompt" required>
            </div>
            <div class="form-group">
                <label for="slides">Number of Slides:</label>
                <input type="number" id="slides" name="slides" value="5" min="1" max="20" required>
            </div>
            <button onclick="generatePresentation()">Generate Presentation</button>
            <div id="result"></div>

            <script>
                function generatePresentation() {
                    const data = {
                        prompt: document.getElementById('prompt').value,
                        slides: parseInt(document.getElementById('slides').value)
                    };

                    // Disable button and show loading state
                    const button = document.querySelector('button');
                    const resultDiv = document.getElementById('result');
                    button.disabled = true;
                    resultDiv.textContent = 'Generating presentation...';
                    resultDiv.className = '';

                    fetch('/generate', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(data)
                    })
                    .then(response => {
                        if (!response.ok) {
                            return response.json().then(err => {
                                throw new Error(err.error || 'Failed to generate presentation');
                            });
                        }
                        return response.json();
                    })
                    .then(data => {
                        const downloadLink = `<br><a href="/download/${data.filename}" class="download-link">Click here to download your presentation</a>`;
                        resultDiv.innerHTML = `<span class="success">${data.result}</span>${downloadLink}`;
                        resultDiv.className = '';
                    })
                    .catch(error => {
                        resultDiv.textContent = 'Error: ' + error.message;
                        resultDiv.className = 'error';
                    })
                    .finally(() => {
                        button.disabled = false;
                    });
                }
            </script>
        </body>
    </html>
    '''

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data provided'}), 400
        
        prompt = data.get('prompt')
        slides = data.get('slides')
        
        if not prompt:
            return jsonify({'error': 'No prompt provided'}), 400
        if not slides:
            return jsonify({'error': 'Number of slides not provided'}), 400
        
        try:
            slides = int(slides)
            if slides < 1 or slides > 20:
                return jsonify({'error': 'Number of slides must be between 1 and 20'}), 400
        except ValueError:
            return jsonify({'error': 'Invalid number of slides'}), 400

        result, filename = generate_ppt(
            prompt,
            'openai',
            'gpt-3.5-turbo',
            slides
        )
        return jsonify({
            'result': 'Your presentation has been generated successfully!',
            'filename': filename
        })
    except Exception as e:
        app.logger.error(f"Error generating presentation: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        app.logger.error(f"Error downloading file: {str(e)}", exc_info=True)
        return jsonify({'error': 'Error downloading file'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
