from flask import Flask, request, jsonify
import os
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
                input { width: 100%; padding: 8px; margin-bottom: 10px; }
                button { background: #4CAF50; color: white; padding: 10px 15px; border: none; cursor: pointer; }
                button:hover { background: #45a049; }
                #result { margin-top: 20px; padding: 10px; }
            </style>
        </head>
        <body>
            <h1>Decklyst - Generate Professional Presentations</h1>
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
                        slides: document.getElementById('slides').value
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

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    result = generate_ppt(
        data['prompt'],
        'openai',  
        'gpt-3.5-turbo',  
        data['slides']
    )
    return jsonify({'result': result})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
