from datetime import datetime, timedelta
from flask_sqlalchemy import SQLAlchemy
from enum import Enum

db = SQLAlchemy()

class PlanType(Enum):
    FREE = 'free'
    PAY_PER_PRESENTATION = 'pay_per_presentation'
    SUBSCRIPTION = 'subscription'

class User(db.Model):
    id = db.Column(db.String(128), primary_key=True)  # Google OAuth ID
    email = db.Column(db.String(120), unique=True, nullable=False)
    name = db.Column(db.String(120))
    picture = db.Column(db.String(500))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    current_plan = db.Column(db.Enum(PlanType), default=PlanType.FREE)
    subscription_start = db.Column(db.DateTime)
    subscription_end = db.Column(db.DateTime)
    monthly_presentations_count = db.Column(db.Integer, default=0)
    monthly_slides_count = db.Column(db.Integer, default=0)
    last_count_reset = db.Column(db.DateTime, default=datetime.utcnow)

    presentations = db.relationship('Presentation', backref='user', lazy=True)

    def can_create_presentation(self, num_slides):
        if self.current_plan == PlanType.FREE:
            active_presentations = Presentation.query.filter_by(
                user_id=self.id,
                status='active'
            ).count()
            return active_presentations < 3 and num_slides <= 5

        elif self.current_plan == PlanType.PAY_PER_PRESENTATION:
            return num_slides <= 10

        elif self.current_plan == PlanType.SUBSCRIPTION:
            # Check if we need to reset monthly counters
            if (datetime.utcnow() - self.last_count_reset).days >= 30:
                self.monthly_presentations_count = 0
                self.monthly_slides_count = 0
                self.last_count_reset = datetime.utcnow()
                db.session.commit()

            return (self.monthly_presentations_count < 50 and 
                    self.monthly_slides_count + num_slides <= 500)

        return False

    def increment_usage(self, num_slides):
        if self.current_plan == PlanType.SUBSCRIPTION:
            self.monthly_presentations_count += 1
            self.monthly_slides_count += num_slides
            db.session.commit()

class Presentation(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.String(128), db.ForeignKey('user.id'), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    num_slides = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    expires_at = db.Column(db.DateTime)
    status = db.Column(db.String(20), default='active')  # active, expired, archived
    file_path = db.Column(db.String(500))

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set expiration for free plan presentations
        user = User.query.get(self.user_id)
        if user and user.current_plan == PlanType.FREE:
            self.expires_at = datetime.utcnow() + timedelta(days=7)

class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.String(128), db.ForeignKey('user.id'), nullable=False)
    amount = db.Column(db.Float, nullable=False)
    payment_type = db.Column(db.String(50), nullable=False)  # one_time, subscription
    paystack_reference = db.Column(db.String(100), unique=True)
    status = db.Column(db.String(20), default='pending')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    subscription_id = db.Column(db.String(100))  # For subscription payments
