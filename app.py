import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session, jsonify
from werkzeug.utils import secure_filename
from generate_ppt import generate_ppt
from auth.google_auth import GoogleAuth, login_required

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Required for session management
auth = GoogleAuth(app)

UPLOAD_FOLDER = 'generated_presentations'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    if 'user' in session:
        return render_template('index.html', user=session['user'])
    return redirect(url_for('login'))

@app.route('/login')
def login():
    if 'user' in session:
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/login/google')
def google_login():
    """Initiate Google OAuth flow"""
    return redirect(auth.get_auth_url())

@app.route('/login/google/authorized')
def google_auth_callback():
    """Handle Google OAuth callback"""
    if 'error' in request.args:
        return redirect(url_for('login'))
    
    if 'code' not in request.args:
        return redirect(url_for('login'))
    
    # Exchange code for tokens
    token_data = auth.get_token(request.args.get('code'))
    if 'access_token' not in token_data:
        return redirect(url_for('login'))
    
    # Get user info
    user_info = auth.get_user_info(token_data['access_token'])
    
    # Store user info in session
    session['user'] = {
        'id': user_info['id'],
        'email': user_info['email'],
        'name': user_info.get('name', ''),
        'picture': user_info.get('picture', '')
    }
    
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

@app.route('/generate', methods=['POST'])
@login_required
def generate():
    """Generate presentation endpoint"""
    topic = request.form.get('topic')
    num_slides = int(request.form.get('num_slides', 5))
    theme = request.form.get('theme', 'professional')
    
    if not topic:
        return jsonify({'error': 'Topic is required'}), 400
        
    try:
        filename = generate_ppt(topic, num_slides, theme)
        return jsonify({
            'success': True,
            'filename': filename,
            'download_url': url_for('download', filename=filename)
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
@login_required
def download(filename):
    """Download generated presentation"""
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
