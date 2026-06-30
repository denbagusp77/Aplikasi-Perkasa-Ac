from app import app, db

with app.app_context():
    print("Database URL:", db.engine.url)
