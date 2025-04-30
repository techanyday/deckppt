import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session, jsonify
from werkzeug.utils import secure_filename
from auth.google_auth import GoogleAuth, login_required
from models.database import db, User, Presentation, PlanType
from services.paystack import PaystackService
from datetime import datetime, timedelta
import logging
from sqlalchemy import text
import tempfile
import shutil
from slides_generator import GoogleSlidesGenerator
import secrets

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', os.urandom(24))

# Configure SQLAlchemy
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///app.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Only add pooling options for PostgreSQL
if 'postgres' in app.config['SQLALCHEMY_DATABASE_URI']:
    app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {
        'pool_size': 5,
        'pool_recycle': 280,
        'pool_timeout': 20,
        'max_overflow': 2
    }
db.init_app(app)

# Initialize services
auth = GoogleAuth(app)
if os.environ.get('PAYSTACK_SECRET_KEY'):
    paystack = PaystackService()
else:
    paystack = None  # Skip Paystack for local development
if os.environ.get('GOOGLE_SLIDES_CREDENTIALS'):
    slides = GoogleSlidesGenerator()  # Will use env vars
else:
    slides = GoogleSlidesGenerator('slides_credentials.json')  # Will use local file

# Configure upload and download directories
UPLOAD_FOLDER = os.path.join('static', 'downloads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Create directory if it doesn't exist
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Create database tables
with app.app_context():
    db.create_all()

@app.before_request
def before_request():
    """Ensure database connection is active"""
    try:
        db.session.execute(text('SELECT 1'))
    except Exception as e:
        logger.error(f"Database connection error: {e}")
        db.session.rollback()
        return jsonify({'error': 'Service temporarily unavailable'}), 503

@app.errorhandler(500)
def internal_error(error):
    """Handle internal server errors"""
    logger.error(f"Internal Server Error: {error}")
    db.session.rollback()
    return jsonify({'error': 'An unexpected error occurred'}), 500

@app.errorhandler(404)
def not_found_error(error):
    """Handle not found errors"""
    return jsonify({'error': 'Resource not found'}), 404

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
def generate_presentation():
    try:
        # Get form data
        topic = request.form.get('title', '').strip()  # Using 'title' as both title and topic
        num_slides = request.form.get('num_slides', '5')
        
        # Validate input
        if not topic:
            return jsonify({
                "success": False,
                "message": "Please provide a topic for your presentation"
            }), 400
            
        try:
            num_slides = int(num_slides)
            if num_slides < 1 or num_slides > 10:
                return jsonify({
                    "success": False,
                    "message": "Number of slides must be between 1 and 10"
                }), 400
        except ValueError:
            return jsonify({
                "success": False,
                "message": "Invalid number of slides"
            }), 400
        
        # Initialize slides service with stored credentials
        if 'slides_credentials' not in session:
            return jsonify({
                "success": False,
                "message": "Please authenticate with Google first"
            }), 401

        # Initialize slides generator
        generator = GoogleSlidesGenerator()
        generator.init_service()
        
        # Create presentation with topic as both title and content topic
        presentation_id = generator.create_presentation(
            title=topic,  # Use topic as title
            topic=topic,  # Use same topic for content generation
            num_slides=num_slides
        )
        
        if not presentation_id:
            return jsonify({
                "success": False,
                "message": "Failed to create presentation. Please try again."
            }), 500
            
        # Get presentation URL
        presentation_url = f'https://docs.google.com/presentation/d/{presentation_id}'
        
        return jsonify({
            "success": True,
            "message": "Presentation created successfully!",
            "url": presentation_url
        })
        
    except Exception as e:
        app.logger.error(f"Error generating presentation: {str(e)}")
        return jsonify({
            "success": False,
            "message": "An error occurred while creating your presentation"
        }), 500

@app.route('/oauth2callback')
def oauth2callback():
    try:
        # Get the authorization code from the callback
        code = request.args.get('code')
        state = request.args.get('state')
        
        if not code:
            app.logger.error("No code received in callback")
            return jsonify({"error": "No authorization code received"}), 400
            
        # Get credentials from the authorization code
        try:
            credentials = slides.get_credentials_from_code(code, state)
            
            # Store credentials in session
            session['slides_credentials'] = {
                'token': credentials.token,
                'refresh_token': credentials.refresh_token,
                'token_uri': credentials.token_uri,
                'client_id': credentials.client_id,
                'client_secret': credentials.client_secret,
                'scopes': credentials.scopes
            }
            
            return redirect(url_for('index'))
            
        except Exception as e:
            app.logger.error(f"Error getting credentials: {str(e)}")
            # If there's a scope mismatch, try to get a new authorization URL
            if "Scope has changed" in str(e):
                auth_url, state = slides.get_authorization_url()
                session['oauth_state'] = state
                return redirect(auth_url)
            raise
            
    except Exception as e:
        app.logger.error(f"OAuth callback error: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/download/<filename>')
@login_required
def download(filename):
    """Download a presentation file."""
    try:
        # Ensure the filename is secure
        filename = secure_filename(filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        
        if not os.path.exists(file_path):
            logger.error(f"File not found: {file_path}")
            return jsonify({'error': 'File not found'}), 404
            
        return send_from_directory(
            UPLOAD_FOLDER,
            filename,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        return jsonify({'error': 'Error downloading file'}), 500

@app.route('/pricing')
@login_required
def pricing():
    user = User.query.get(session['user']['id'])
    return render_template('pricing.html', user=user)

@app.route('/payment/create', methods=['POST'])
@login_required
def create_payment():
    try:
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
        
        return jsonify({
            'error': 'Payment initialization failed. Please contact support if this persists.'
        }), 500
    except ValueError as e:
        # This will catch the missing API key error
        return jsonify({
            'error': 'Payment system is not properly configured. Please contact support.'
        }), 500
    except Exception as e:
        return jsonify({
            'error': f'An unexpected error occurred: {str(e)}'
        }), 500

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
    """Enhanced health check endpoint"""
    try:
        # Check database connection
        db.session.execute(text('SELECT 1'))
        
        # Check file system
        if not os.path.exists(UPLOAD_FOLDER):
            raise Exception("Upload directory not accessible")
            
        # Check if we can write to the upload directory
        test_file = os.path.join(UPLOAD_FOLDER, '.health_check')
        try:
            with open(test_file, 'w') as f:
                f.write('ok')
            os.remove(test_file)
        except Exception as e:
            raise Exception(f"Upload directory not writable: {e}")
        
        return jsonify({
            'status': 'healthy',
            'timestamp': datetime.utcnow().isoformat(),
            'checks': {
                'database': 'ok',
                'filesystem': 'ok'
            }
        }), 200
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        return jsonify({
            'status': 'unhealthy',
            'error': str(e),
            'timestamp': datetime.utcnow().isoformat()
        }), 503

@app.route('/privacy')
def privacy_policy():
    return render_template('privacy.html', now=datetime.now())

@app.route('/terms')
def terms_of_service():
    return render_template('terms.html', now=datetime.now())

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
