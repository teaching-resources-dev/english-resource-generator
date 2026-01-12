"""
Authentication module for English Resource Generator.

Simple email allowlist authentication - no passwords required.
Approved emails are stored in Streamlit secrets.
"""

import streamlit as st
from typing import List
import hashlib
import time


def get_allowed_emails() -> List[str]:
    """Get list of allowed emails from secrets."""
    try:
        return st.secrets.auth.allowed_emails
    except (KeyError, AttributeError):
        # For local development without secrets
        st.warning("No allowed_emails configured in secrets. Using demo mode.")
        return ["demo@example.com"]


def get_cookie_key() -> str:
    """Get cookie key from secrets."""
    try:
        return st.secrets.auth.cookie_key
    except (KeyError, AttributeError):
        return "default-dev-key-change-in-production"


def check_authentication() -> bool:
    """
    Check if the current user is authenticated.

    Returns:
        True if user is authenticated, False otherwise.
    """
    return st.session_state.get("authenticated", False)


def authenticate_user(email: str) -> bool:
    """
    Attempt to authenticate a user by email.

    Args:
        email: The email address to check.

    Returns:
        True if authentication successful, False otherwise.
    """
    email = email.strip().lower()
    allowed_emails = [e.lower() for e in get_allowed_emails()]

    if email in allowed_emails:
        st.session_state.authenticated = True
        st.session_state.user_email = email
        st.session_state.login_time = time.time()
        return True

    return False


def show_login_page():
    """Display the login page."""
    st.title("📚 English Resource Generator")
    st.markdown("### Teacher Login")
    st.markdown("Enter your school email to access resource generation.")

    with st.form("login_form"):
        email = st.text_input(
            "Email Address",
            placeholder="your.name@school.wa.edu.au",
            help="Use your authorised school email address"
        )

        submitted = st.form_submit_button("Login", use_container_width=True)

        if submitted:
            if not email:
                st.error("Please enter your email address.")
            elif authenticate_user(email):
                st.success("Login successful!")
                st.rerun()
            else:
                st.error(
                    "Email not authorised. Contact your administrator "
                    "if you believe this is an error."
                )

    st.markdown("---")
    st.caption(
        "This resource is for authorised staff only. "
        "Your email must be on the approved list to access."
    )
