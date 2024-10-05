import json
from google.oauth2 import service_account
from google.auth.transport import requests
from flask import jsonify, request

def create_signed_jwt(request):
    request_json = request.get_json(silent=True)
    
    if not request_json or 'service_account_info' not in request_json:
        return jsonify({'error': 'Invalid request: service_account_info is required'}), 400
    
    service_account_info = request_json['service_account_info']
    
    try:
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info,
            scopes=['https://www.googleapis.com/auth/bigquery']
        )
        
        auth_req = requests.Request()
        credentials.refresh(auth_req)
        jwt_token = credentials.token
        
        return jsonify({'jwt': jwt_token})
    except ValueError as ve:
        return jsonify({'error': f'Invalid service account info: {str(ve)}'}), 400
    except Exception as e:
        return jsonify({'error': f'An error occurred: {str(e)}'}), 500