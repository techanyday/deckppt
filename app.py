from flask import Flask, request, jsonify
import utils
from generate_ppt import generate_ppt

app = Flask(__name__)

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
                input, select { width: 100%; padding: 8px; margin-bottom: 10px; }
                button { background: #4CAF50; color: white; padding: 10px 15px; border: none; cursor: pointer; }
                button:hover { background: #45a049; }
            </style>
        </head>
        <body>
            <h1>Decklyst - Generate Professional Presentations</h1>
            <div class="form-group">
                <label for="prompt">Prompt (write me a PPT presentation about...):</label>
                <input type="text" id="prompt" name="prompt" required>
            </div>
            <div class="form-group">
                <label for="api">Select Generation API:</label>
                <select id="api" name="api" required>
                    <option value="">Select an option</option>
                </select>
            </div>
            <div class="form-group">
                <label for="model">Select Script Generation Model:</label>
                <select id="model" name="model" required>
                    <option value="">Select an API first</option>
                </select>
            </div>
            <div class="form-group">
                <label for="slides">Number of Slides:</label>
                <input type="number" id="slides" name="slides" required>
            </div>
            <div class="form-group">
                <label for="api_key">API Key:</label>
                <input type="password" id="api_key" name="api_key" required>
            </div>
            <button onclick="generatePresentation()">Generate Presentation</button>
            <div id="result"></div>

            <script>
                // Load API options on page load
                fetch('/api/options')
                    .then(response => response.json())
                    .then(data => {
                        const apiSelect = document.getElementById('api');
                        data.forEach(api => {
                            const option = document.createElement('option');
                            option.value = api;
                            option.textContent = api;
                            apiSelect.appendChild(option);
                        });
                    });

                // Update model options when API is selected
                document.getElementById('api').addEventListener('change', function() {
                    fetch(`/api/models/${this.value}`)
                        .then(response => response.json())
                        .then(data => {
                            const modelSelect = document.getElementById('model');
                            modelSelect.innerHTML = '';
                            data.forEach(model => {
                                const option = document.createElement('option');
                                option.value = model;
                                option.textContent = model;
                                modelSelect.appendChild(option);
                            });
                        });
                });

                function generatePresentation() {
                    const data = {
                        prompt: document.getElementById('prompt').value,
                        api: document.getElementById('api').value,
                        model: document.getElementById('model').value,
                        slides: document.getElementById('slides').value,
                        api_key: document.getElementById('api_key').value
                    };

                    fetch('/generate', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify(data)
                    })
                    .then(response => response.json())
                    .then(data => {
                        document.getElementById('result').textContent = data.result;
                    })
                    .catch(error => {
                        document.getElementById('result').textContent = 'Error: ' + error.message;
                    });
                }
            </script>
        </body>
    </html>
    '''

@app.route('/api/options')
def api_options():
    return jsonify(utils.get_api_list())

@app.route('/api/models/<api_name>')
def model_options(api_name):
    return jsonify(utils.get_model_list_from_api(api_name))

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    utils.save_config(data['api_key'])
    result = generate_ppt(
        data['prompt'],
        data['api'],
        data['model'],
        data['slides']
    )
    return jsonify({'result': result})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
