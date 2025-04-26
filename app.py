from flask import Flask, request, jsonify, send_file
from generate_ppt import generate_ppt
import os
import logging

app = Flask(__name__)

UPLOAD_FOLDER = 'generated_presentations'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>DeckSky - AI Presentation Generator</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body { padding: 20px; background-color: #f8f9fa; }
            .container { max-width: 800px; }
            .card { border-radius: 15px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); }
            .theme-select { margin-bottom: 20px; }
            .btn-primary { background-color: #0d6efd; border-color: #0d6efd; }
            .btn-primary:hover { background-color: #0b5ed7; border-color: #0b5ed7; }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="card p-4 mb-4">
                <h1 class="text-center mb-4">DeckSky</h1>
                <h5 class="text-center text-muted mb-4">AI-Powered Presentation Generator</h5>
                
                <form id="generateForm" class="needs-validation" novalidate>
                    <div class="mb-3">
                        <label for="topic" class="form-label">Topic</label>
                        <input type="text" class="form-control" id="topic" required
                               placeholder="Enter your presentation topic">
                        <div class="form-text">Be specific and descriptive for better results</div>
                    </div>

                    <div class="mb-3">
                        <label for="theme" class="form-label">Theme</label>
                        <select class="form-select" id="theme" required>
                            <option value="professional">Professional (Clean & Business-like)</option>
                            <option value="modern">Modern (Contemporary Design)</option>
                            <option value="minimal">Minimal (Simple & Elegant)</option>
                            <option value="creative">Creative (Bold & Artistic)</option>
                            <option value="corporate">Corporate (Formal & Structured)</option>
                        </select>
                        <div class="form-text">Choose a theme that matches your presentation style</div>
                    </div>

                    <div class="mb-3">
                        <label for="slides" class="form-label">Number of Slides</label>
                        <input type="number" class="form-control" id="slides" 
                               min="3" max="15" value="5" required>
                        <div class="form-text">Recommended: 5-10 slides for optimal content</div>
                    </div>

                    <div class="row">
                        <div class="col-md-6 mb-3">
                            <label for="api" class="form-label">AI Model Provider</label>
                            <select class="form-select" id="api" required>
                                <option value="openai">OpenAI</option>
                                <option value="huggingface">HuggingFace</option>
                            </select>
                        </div>

                        <div class="col-md-6 mb-3">
                            <label for="model" class="form-label">Model</label>
                            <select class="form-select" id="model" required>
                                <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
                                <option value="gpt-4">GPT-4</option>
                                <option value="flan-t5-small">Flan-T5-Small</option>
                            </select>
                        </div>
                    </div>

                    <div class="d-grid gap-2">
                        <button type="submit" class="btn btn-primary btn-lg">Generate Presentation</button>
                    </div>
                </form>
            </div>

            <div id="status" class="alert mt-3" style="display: none;"></div>
        </div>

        <script>
            document.getElementById('generateForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                const status = document.getElementById('status');
                const submitButton = e.target.querySelector('button[type="submit"]');
                
                status.style.display = 'block';
                status.className = 'alert alert-info';
                status.textContent = 'Generating your presentation...';
                submitButton.disabled = true;

                const data = {
                    topic: document.getElementById('topic').value,
                    theme: document.getElementById('theme').value,
                    num_slides: parseInt(document.getElementById('slides').value),
                    api_name: document.getElementById('api').value,
                    model_name: document.getElementById('model').value
                };

                try {
                    const response = await fetch('/generate', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify(data)
                    });

                    const result = await response.json();
                    if (result.success) {
                        status.className = 'alert alert-success';
                        status.textContent = 'Presentation generated successfully! Downloading...';
                        window.location.href = `/download/${result.filename}`;
                    } else {
                        status.className = 'alert alert-danger';
                        status.textContent = result.error || 'Failed to generate presentation';
                    }
                } catch (error) {
                    status.className = 'alert alert-danger';
                    status.textContent = 'Error generating presentation';
                } finally {
                    submitButton.disabled = false;
                }
            });

            // Update model options based on selected API
            document.getElementById('api').addEventListener('change', (e) => {
                const modelSelect = document.getElementById('model');
                const options = e.target.value === 'openai' 
                    ? [
                        ['gpt-3.5-turbo', 'GPT-3.5 Turbo'],
                        ['gpt-4', 'GPT-4']
                      ]
                    : [
                        ['flan-t5-small', 'Flan-T5-Small']
                      ];
                
                modelSelect.innerHTML = options
                    .map(([value, text]) => `<option value="${value}">${text}</option>`)
                    .join('');
            });
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

        topic = data.get('topic')
        theme = data.get('theme', 'professional')
        num_slides = data.get('num_slides', 5)
        api_name = data.get('api_name', 'openai')
        model_name = data.get('model_name', 'gpt-3.5-turbo')

        if not topic:
            return jsonify({'error': 'Topic is required'}), 400

        filename = generate_ppt(topic, api_name, model_name, num_slides, theme)
        return jsonify({'success': True, 'filename': filename})

    except Exception as e:
        logging.error(f"Error generating presentation: {str(e)}", exc_info=True)
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        logging.error(f"Error downloading file: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
