
# App configuration
import os

class Settings:
    def __init__(self):
        self.debug = os.getenv('GECE_DEBUG', '0') == '1'
        self.db_path = os.getenv('GECE_DB_PATH', '')

settings = Settings()
