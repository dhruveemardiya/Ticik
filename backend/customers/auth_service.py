import hashlib
import json
import secrets
import datetime
from pathlib import Path
from django.conf import settings

class AuthService:
    _instance = None

    def __init__(self):
        self.auth_file = Path(settings.BASE_DIR) / 'media' / 'customer_data' / 'auth_config.json'
        self.auth_file.parent.mkdir(parents=True, exist_ok=True)
        self.active_tokens = set()
        self._ensure_config()

    @classmethod
    def get_instance(cls):
        if cls._instance is None:
            cls._instance = AuthService()
        return cls._instance

    def _hash_password(self, password: str, salt: bytes = None) -> str:
        if salt is None:
            salt = secrets.token_bytes(16)
        key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000)
        return f"{salt.hex()}:{key.hex()}"

    def _verify_hash(self, password: str, stored_hash: str) -> bool:
        try:
            salt_hex, key_hex = stored_hash.split(':')
            salt = bytes.fromhex(salt_hex)
            key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 100000)
            return secrets.compare_digest(key.hex(), key_hex)
        except Exception:
            return False

    def _ensure_config(self):
        if not self.auth_file.exists():
            initial_data = {
                "password_hash": self._hash_password("Trip@2026"),
                "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "updated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            with open(self.auth_file, 'w', encoding='utf-8') as f:
                json.dump(initial_data, f, indent=2)

    def _load_config(self):
        self._ensure_config()
        try:
            with open(self.auth_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {"password_hash": self._hash_password("Trip@2026")}

    def check_password(self, password: str) -> bool:
        if not password:
            return False
        config = self._load_config()
        stored_hash = config.get("password_hash", "")
        return self._verify_hash(password, stored_hash)

    def create_token(self) -> str:
        token = secrets.token_hex(24)
        self.active_tokens.add(token)
        return token

    def verify_token(self, token: str) -> bool:
        return bool(token and token in self.active_tokens)

    def revoke_token(self, token: str):
        if token in self.active_tokens:
            self.active_tokens.remove(token)
