import json
import secrets
import requests
from time import time


class OfficeUser:

    def __init__(self, client_id: str, tenant_id: str, client_secret: str):
        self._client_id = client_id
        self._client_secret = client_secret
        self._tenant_id = tenant_id
        self._token = None
        self._token_expires = None

    @staticmethod
    def _password_gen():
        return secrets.token_urlsafe(8)

    def _refresh_token(self):
        if self._token is None or \
                time() > self._token_expires:
            self._get_token()

    def _get_token(self):
        r = requests.post(
            url=f'https://login.microsoftonline.com/{self._tenant_id}/oauth2/v2.0/token',
            headers={
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            data={
                'grant_type': 'client_credentials',
                'client_id': self._client_id,
                'client_secret': self._client_secret,
                'scope': 'https://graph.microsoft.com/.default'
            }
        )
        data = r.json()

        if r.status_code != 200 or \
                'error' in data:
            raise Exception(json.dumps(data))

        self._token = data['access_token']
        self._token_expires = int(data['expires_in']) + int(time())

    def _assign_license(self, email: str, sku_id: str):
        self._refresh_token()

        r = requests.post(
            url=f'https://graph.microsoft.com/v1.0/users/{email}/assignLicense',
            headers={
                'Authorization': f'Bearer {self._token}',
                'Content-Type': 'application/json'
            },
            json={
                'addLicenses': [{
                    'disabledPlans': [],
                    'skuId': sku_id
                }],
                'removeLicenses': []
            }
        )
        data = r.json()

        if r.status_code != 200 or \
                'error' in data:
            raise Exception(json.dumps(data))

    def _create_user(
            self,
            display_name: str,
            username: str,
            password: str,
            domain: str,
            location: str = 'CN'
    ):
        self._refresh_token()

        r = requests.post(
            url='https://graph.microsoft.com/v1.0/users',
            headers={
                'Authorization': f'Bearer {self._token}',
                'Content-Type': 'application/json'
            },
            json={
                'accountEnabled': True,
                'displayName': display_name,
                'mailNickname': username,
                'passwordPolicies': 'DisablePasswordExpiration, DisableStrongPassword',
                'passwordProfile': {
                    'password': password,
                    'forceChangePasswordNextSignIn': True
                },
                'userPrincipalName': username + domain,
                'usageLocation': location
            }
        )
        data = r.json()

        if r.status_code != 201 or \
                'error' in data:
            raise Exception(json.dumps(data))

    def create_account(
            self,
            username: str,
            domain: str,
            sku_id: str,
            display_name: str = None,
            password: str = None,
            location: str = 'CN'
    ):
        if display_name is None:
            display_name = username

        if password is None:
            password = self._password_gen()

        email = username + domain

        self._create_user(
            display_name=display_name,
            username=username,
            password=password,
            domain=domain,
            location=location
        )
        self._assign_license(
            email=email,
            sku_id=sku_id
        )

        return {
            'email': email,
            'password': password
        }
