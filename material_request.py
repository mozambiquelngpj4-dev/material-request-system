
from datetime import datetime

from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)


db = SQLAlchemy()


class MaterialRequest(db.Model):

    __tablename__ = "material_requests"

    id = db.Column(
        db.BigInteger,
        primary_key=True
    )

    sn = db.Column(
        db.Integer,
        nullable=True
    )

    company = db.Column(
        db.String(200),
        nullable=False
    )

    mri_no = db.Column(
        db.String(100),
        nullable=False,
        unique=True
    )

    requested_by = db.Column(
        db.String(200),
        nullable=False
    )

    req_date = db.Column(
        db.Date,
        nullable=False
    )

    ltsc_status = db.Column(
        db.String(50),
        nullable=False,
        default="Pending"
    )

    piping_status = db.Column(
        db.String(50),
        nullable=False,
        default="Pending"
    )

    piping_date = db.Column(
        db.Date,
        nullable=True
    )

    ccs_jv_confirm = db.Column(
        db.String(50),
        nullable=False,
        default="Pending"
    )

    material_status = db.Column(
        db.String(50),
        nullable=False,
        default="Pending"
    )

    mtrl_date = db.Column(
        db.Date,
        nullable=True
    )

    attachment_excel = db.Column(
        db.String(500),
        nullable=True
    )

    attachment_pdf = db.Column(
        db.String(500),
        nullable=True
    )

    voucher_files = db.Column(
        db.Text,
        nullable=True
    )

    remarks = db.Column(
        db.Text,
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )

    def __repr__(self):
        return f"<MaterialRequest {self.mri_no}>"


class User(UserMixin, db.Model):

    __tablename__ = "users"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    username = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    full_name = db.Column(
        db.String(200),
        nullable=False
    )

    password_hash = db.Column(
        db.String(255),
        nullable=False
    )

    role = db.Column(
        db.String(50),
        nullable=False,
        default="REQUESTER"
    )

    is_active = db.Column(
        db.Boolean,
        default=True,
        nullable=False
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    def set_password(self, password):

        self.password_hash = generate_password_hash(
            password
        )

    def check_password(self, password):

        return check_password_hash(
            self.password_hash,
            password
        )

    def __repr__(self):

        return f"<User {self.username}>"


class ApprovalHistory(db.Model):

    __tablename__ = "approval_history"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    material_request_id = db.Column(
        db.BigInteger,
        db.ForeignKey(
            "material_requests.id"
        ),
        nullable=False
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    stage = db.Column(
        db.String(50),
        nullable=False
    )

    action = db.Column(
        db.String(50),
        nullable=False
    )

    remarks = db.Column(
        db.Text,
        nullable=True
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    material_request = db.relationship(
        "MaterialRequest",
        backref=db.backref(
            "approval_history",
            lazy=True
        )
    )

    user = db.relationship(
        "User",
        backref=db.backref(
            "approval_history",
            lazy=True
        )
    )


# =========================================================
# NOTIFICATIONS
# =========================================================

class Notification(db.Model):

    __tablename__ = "notifications"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    user_id = db.Column(
        db.Integer,
        db.ForeignKey("users.id"),
        nullable=False
    )

    title = db.Column(
        db.String(200),
        nullable=False
    )

    message = db.Column(
        db.Text,
        nullable=False
    )

    notification_type = db.Column(
        db.String(50),
        nullable=False,
        default="INFO"
    )

    material_request_id = db.Column(
        db.BigInteger,
        db.ForeignKey("material_requests.id"),
        nullable=True
    )

    is_read = db.Column(
        db.Boolean,
        nullable=False,
        default=False
    )

    created_at = db.Column(
        db.DateTime,
        default=datetime.utcnow
    )

    user = db.relationship(
        "User",
        backref=db.backref(
            "notifications",
            lazy=True
        )
    )

    material_request = db.relationship(
        "MaterialRequest",
        backref=db.backref(
            "notifications",
            lazy=True
        )
    )

    def __repr__(self):
        return f"<Notification {self.id}>"