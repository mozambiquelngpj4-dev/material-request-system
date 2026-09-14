from app import app
from material_request import db
from sqlalchemy import text


with app.app_context():

    db.session.execute(
        text(
            """
            CREATE TABLE IF NOT EXISTS notifications (
                id SERIAL PRIMARY KEY,

                user_id INTEGER NOT NULL
                    REFERENCES users(id),

                material_request_id BIGINT NULL
                    REFERENCES material_requests(id),

                title VARCHAR(200) NOT NULL,

                message TEXT NOT NULL,

                notification_type VARCHAR(50) NOT NULL
                    DEFAULT 'INFO',

                is_read BOOLEAN NOT NULL
                    DEFAULT FALSE,

                created_at TIMESTAMP
                    DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
    )

    db.session.commit()

    print("Notifications table created successfully.")

