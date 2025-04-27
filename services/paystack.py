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
        try:
            # Check if APP_URL is set
            app_url = os.environ.get("APP_URL")
            if not app_url:
                raise ValueError("APP_URL environment variable is not set")

            # Convert USD to GHS (Paystack expects amount in pesewas - 100 pesewas = 1 GHS)
            amount_ghs = amount_usd * 12.5  # Approximate USD to GHS conversion
            amount_pesewas = int(amount_ghs * 100)

            data = {
                'email': email,
                'amount': amount_pesewas,
                'currency': 'GHS',  # Changed to Ghana Cedis
                'callback_url': f'{app_url}/payment/callback',
                'metadata': {
                    'user_id': user_id,
                    'payment_type': payment_type
                }
            }

            print(f"[Paystack] Using API Key: {self.api_key[:8]}...")
            print(f"[Paystack] Using callback URL: {data['callback_url']}")
            print(f"[Paystack] Initializing transaction with data: {data}")
            
            response = requests.post(
                f'{self.base_url}/transaction/initialize',
                headers=self.headers,
                json=data
            )

            print(f"[Paystack] Response status: {response.status_code}")
            print(f"[Paystack] Response body: {response.text}")

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
            
            # If we get here, something went wrong
            print(f"[Paystack] Error: Non-200 status code")
            return None, None

        except Exception as e:
            print(f"[Paystack] Exception: {str(e)}")
            raise

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
