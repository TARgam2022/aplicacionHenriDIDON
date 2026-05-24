import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from app_server import app

handler = app
