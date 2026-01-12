"""
English Resource Generator - Main Application

A Streamlit web app for generating English teaching resources.
Colleagues can generate quizzes, worked examples, and worksheets
without technical knowledge.
"""

import streamlit as st

# Page configuration - must be first Streamlit command
st.set_page_config(
    page_title="English Resource Generator",
    page_icon="📚",
    layout="centered",
    initial_sidebar_state="expanded"
)

from utils.auth import check_authentication, show_login_page


def main():
    """Main application entry point."""

    # Check if user is authenticated
    if not check_authentication():
        show_login_page()
        return

    # User is authenticated - show main content
    st.title("📚 English Resource Generator")
    st.markdown("Generate teaching resources for your English classes.")

    # Show user info in sidebar
    st.sidebar.success(f"Logged in as: {st.session_state.get('user_email', 'Unknown')}")
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()

    # Main navigation
    st.markdown("---")
    st.subheader("Available Resources")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        ### 📝 Quiz Generator
        Create multiple-choice quizzes with answer keys.

        [Generate Quiz →](./Generate)
        """)

    with col2:
        st.markdown("""
        ### 📖 More Coming Soon
        - Worked Examples
        - Grammar Worksheets
        - 4-Step Analysis
        """)

    # Footer
    st.markdown("---")
    st.caption("English Resource Generator v0.1 | Built for WA English Teachers")


if __name__ == "__main__":
    main()
