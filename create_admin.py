
from getpass import getpass

from app import app
from material_request import db, User


with app.app_context():

    username = input(
        "Username: "
    ).strip()

    full_name = input(
        "Full name: "
    ).strip()

    role = input(
        "Role (ADMIN/REQUESTER/APPROVER_PIPING/APPROVER_MATERIAL): "
    ).strip().upper()

    if role not in [
        "ADMIN",
        "REQUESTER",
        "APPROVER_PIPING",
        "APPROVER_MATERIAL"
    ]:

        print(
            "Invalid role."
        )

        print(
            "Use: ADMIN, REQUESTER, APPROVER_PIPING, "
            "or APPROVER_MATERIAL."
        )

        raise SystemExit

    password = getpass(
        "Password: "
    )

    confirm_password = getpass(
        "Confirm password: "
    )

    if password != confirm_password:

        print(
            "Passwords do not match."
        )

        raise SystemExit

    existing_user = User.query.filter_by(
        username=username
    ).first()

    if existing_user:

        print(
            "Username already exists."
        )

        raise SystemExit

    # Create user without SQLAlchemy constructor keywords
    # to avoid Pylance "No parameter named" warnings.
    user = User()

    user.username = username
    user.full_name = full_name
    user.role = role
    

    user.set_password(
        password
    )

    db.session.add(user)

    db.session.commit()

    print(
        f"{role} user '{username}' created successfully."
    )

