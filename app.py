import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session, jsonify
from werkzeug.utils import secure_filename
from generate_ppt import generate_ppt
from auth.google_auth import GoogleAuth, login_required
from models.database import db, User, Presentation, PlanType
from services.paystack import PaystackService
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', os.urandom(24))

# Configure SQLAlchemy
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///app.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db.init_app(app)

auth = GoogleAuth(app)
paystack = PaystackService()

UPLOAD_FOLDER = 'generated_presentations'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Create database tables
with app.app_context():
    db.create_all()

@app.route('/')
def index():
    if 'user' in session:
        user = User.query.get(session['user']['id'])
        presentations = Presentation.query.filter_by(
            user_id=user.id,
            status='active'
        ).order_by(Presentation.created_at.desc()).all()
        return render_template('index.html', user=user, presentations=presentations)
    return redirect(url_for('login'))

@app.route('/login')
def login():
    if 'user' in session:
        return redirect(url_for('index'))
    return render_template('login.html')

@app.route('/login/google')
def google_login():
    return redirect(auth.get_auth_url())

@app.route('/login/google/authorized')
def google_auth_callback():
    if 'error' in request.args:
        return redirect(url_for('login'))
    
    if 'code' not in request.args:
        return redirect(url_for('login'))
    
    token_data = auth.get_token(request.args.get('code'))
    if 'access_token' not in token_data:
        return redirect(url_for('login'))
    
    user_info = auth.get_user_info(token_data['access_token'])
    
    # Create or update user
    user = User.query.get(user_info['id'])
    if not user:
        user = User(
            id=user_info['id'],
            email=user_info['email'],
            name=user_info.get('name', ''),
            picture=user_info.get('picture', ''),
            current_plan=PlanType.FREE
        )
        db.session.add(user)
        db.session.commit()
    
    session['user'] = {
        'id': user.id,
        'email': user.email,
        'name': user.name,
        'picture': user.picture
    }
    
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

@app.route('/generate', methods=['POST'])
@login_required
def generate():
    user = User.query.get(session['user']['id'])
    topic = request.form.get('topic')
    num_slides = int(request.form.get('num_slides', 5))
    theme = request.form.get('theme', 'professional')
    
    if not topic:
        return jsonify({'error': 'Topic is required'}), 400

    # Check if user can create presentation
    if not user.can_create_presentation(num_slides):
        return jsonify({
            'error': 'Plan limit reached. Please upgrade your plan to continue.',
            'upgrade_url': url_for('pricing'),
            'plan_limit': True
        }), 403
        
    try:
        filename = generate_ppt(topic, num_slides, theme)
        
        # Create presentation record
        presentation = Presentation(
            user_id=user.id,
            title=topic,
            num_slides=num_slides,
            file_path=filename
        )
        db.session.add(presentation)
        
        # Update usage metrics
        user.increment_usage(num_slides)
        
        db.session.commit()
        
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
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

@app.route('/pricing')
@login_required
def pricing():
    user = User.query.get(session['user']['id'])
    return render_template('pricing.html', user=user)

@app.route('/payment/create', methods=['POST'])
@login_required
def create_payment():
    user = User.query.get(session['user']['id'])
    payment_type = request.form.get('type')
    
    if payment_type == 'subscription':
        authorization_url, reference = paystack.create_subscription(user.id, user.email)
    else:  # one_time
        authorization_url, reference = paystack.initialize_transaction(user.id, user.email, 0.99)
    
    if authorization_url:
        return jsonify({
            'success': True,
            'authorization_url': authorization_url
        })
    
    return jsonify({'error': 'Payment initialization failed'}), 500

@app.route('/payment/callback')
def payment_callback():
    reference = request.args.get('reference')
    if reference and paystack.verify_transaction(reference):
        return redirect(url_for('index'))
    return redirect(url_for('pricing'))

@app.route('/subscription/cancel')
@login_required
def cancel_subscription():
    user = User.query.get(session['user']['id'])
    if paystack.cancel_subscription(user.id):
        return redirect(url_for('pricing'))
    return jsonify({'error': 'Failed to cancel subscription'}), 500

@app.route('/healthz')
def health_check():
    return jsonify({'status': 'healthy'}), 200

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
