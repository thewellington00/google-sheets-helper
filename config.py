import os
from dotenv import load_dotenv

load_dotenv()

GOOGLE_SERVICE_ACCOUNT_KEY = os.getenv('GOOGLE_SERVICE_ACCOUNT_KEY')
def get_google_service_account_key():
    return GOOGLE_SERVICE_ACCOUNT_KEY