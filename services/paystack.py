import os
import requests
from datetime import datetime, timedelta
from models.database import db, User, Payment, PlanType

class PaystackService:
    def __init__(self):
        self.api_key = os.environ.get('PAYSTACK_SECRET_KEY')
        if not self.api_key:
            raise ValueError('PAYSTACK_SECRET_KEY environment variable is not set')
        self.base_url = 'https://api.paystack.co'
        self.headers = {
            'Authorization': f'Bearer {self.api_key}',
            'Content-Type': 'application/json'
        }

    def initialize_transaction(self, user_id, email, amount_usd, payment_type='one_time'):
        # Convert USD to NGN (Paystack uses kobo - 100 kobo = 1 NGN)
        amount_ngn = amount_usd * 750  # Approximate USD to NGN conversion
        amount_kobo = int(amount_ngn * 100)

        data = {
            'email': email,
            'amount': amount_kobo,
            'currency': 'NGN',
            'callback_url': f'{os.environ.get("APP_URL")}/payment/callback',
            'metadata': {
                'user_id': user_id,
                'payment_type': payment_type
            }
        }

        response = requests.post(
            f'{self.base_url}/transaction/initialize',
            headers=self.headers,
            json=data
        )

        if response.status_code == 200:
            result = response.json()
            # Create payment record
            payment = Payment(
                user_id=user_id,
                amount=amount_usd,
                payment_type=payment_type,
                paystack_reference=result['data']['reference']
            )
            db.session.add(payment)
            db.session.commit()

            return result['data']['authorization_url'], result['data']['reference']
        
        return None, None

    def verify_transaction(self, reference):
        response = requests.get(
            f'{self.base_url}/transaction/verify/{reference}',
            headers=self.headers
        )

        if response.status_code == 200:
            result = response.json()
            if result['data']['status'] == 'success':
                # Update payment status
                payment = Payment.query.filter_by(paystack_reference=reference).first()
                if payment:
                    payment.status = 'success'
                    
                    # Update user's plan
                    user = User.query.get(payment.user_id)
                    if payment.payment_type == 'subscription':
                        user.current_plan = PlanType.SUBSCRIPTION
                        user.subscription_start = datetime.utcnow()
                        user.subscription_end = datetime.utcnow() + timedelta(days=30)
                    else:  # one_time payment
                        user.current_plan = PlanType.PAY_PER_PRESENTATION

                    db.session.commit()
                    return True

        return False

    def create_subscription(self, user_id, email):
        # Initialize transaction for subscription
        return self.initialize_transaction(user_id, email, 2.99, 'subscription')

    def cancel_subscription(self, user_id):
        user = User.query.get(user_id)
        if user and user.current_plan == PlanType.SUBSCRIPTION:
            user.current_plan = PlanType.FREE
            user.subscription_end = datetime.utcnow()
            db.session.commit()
            return True
        return False
