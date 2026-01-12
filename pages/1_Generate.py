"""
Generate Resource Page

Main page for generating teaching resources.
"""

import json
from pathlib import Path

import streamlit as st

from utils.auth import check_authentication
from generators.quiz_generator import generate_quiz_docx, load_text_index
from generators.llm_client import RateLimitError, ContentFilterError, GenerationError


# Page config
st.set_page_config(
    page_title="Generate Resource",
    page_icon="📝",
    layout="centered"
)


def main():
    # Check authentication
    if not check_authentication():
        st.warning("Please log in from the home page to access this feature.")
        st.switch_page("app.py")
        return

    st.title("📝 Generate Resource")

    # Load text index
    index = load_text_index()

    # Resource type selection
    st.subheader("1. Select Resource Type")
    resource_type = st.selectbox(
        "What would you like to create?",
        options=["Quiz"],  # More options will be added in Phase 2
        help="Select the type of resource to generate"
    )

    # Year level selection
    st.subheader("2. Select Year Level and Text")

    col1, col2 = st.columns(2)

    with col1:
        year_options = {yl["id"]: yl["label"] for yl in index["year_levels"]}
        selected_year = st.selectbox(
            "Year Level",
            options=list(year_options.keys()),
            format_func=lambda x: year_options[x]
        )

    # Filter texts by year level
    available_texts = [t for t in index["texts"] if t["year"] == selected_year]

    with col2:
        if available_texts:
            text_options = {t["id"]: f"{t['name']} ({t['type']})" for t in available_texts}
            selected_text = st.selectbox(
                "Text",
                options=list(text_options.keys()),
                format_func=lambda x: text_options[x]
            )
        else:
            st.warning("No texts available for this year level.")
            selected_text = None

    # Quiz-specific options
    if resource_type == "Quiz" and selected_text:
        st.subheader("3. Quiz Details")

        topic = st.text_input(
            "Topic",
            placeholder="e.g., Character analysis, Themes of conformity, Language features",
            help="What should the quiz focus on?"
        )

        num_questions = st.slider(
            "Number of Questions",
            min_value=5,
            max_value=15,
            value=10,
            help="Recommended: 10 questions for a standard quiz"
        )

        # Get text info for display
        text_info = next((t for t in available_texts if t["id"] == selected_text), None)

        if text_info:
            with st.expander("Available Knowledge Base Resources"):
                st.write(f"**Text**: {text_info['name']}")
                st.write(f"**Type**: {text_info['type']}")
                st.write(f"**Resources**: {', '.join(text_info.get('resources', ['None']))}")

        # Generate button
        st.subheader("4. Generate")

        if st.button("🚀 Generate Quiz", type="primary", use_container_width=True):
            if not topic:
                st.error("Please enter a topic for the quiz.")
            else:
                with st.spinner("Generating quiz... This may take 30-60 seconds."):
                    try:
                        docx_buffer, raw_content = generate_quiz_docx(
                            year_level=selected_year.replace("F", " Fundamentals"),
                            text_id=selected_text,
                            topic=topic,
                            num_questions=num_questions
                        )

                        # Store in session state
                        st.session_state.generated_docx = docx_buffer
                        st.session_state.generated_content = raw_content
                        st.session_state.generated_filename = (
                            f"Year{selected_year}_{text_info['name'].replace(' ', '_')}"
                            f"_Quiz_{topic.replace(' ', '_')}.docx"
                        )

                        st.success("✅ Quiz generated successfully!")

                    except RateLimitError as e:
                        st.error(f"⏳ {str(e)}")
                    except ContentFilterError as e:
                        st.error(f"🚫 {str(e)}")
                    except GenerationError as e:
                        st.error(f"❌ {str(e)}")
                    except Exception as e:
                        st.error(f"❌ An unexpected error occurred: {str(e)}")

        # Show download and preview if content was generated
        if st.session_state.get("generated_docx"):
            st.divider()

            col1, col2 = st.columns(2)

            with col1:
                st.download_button(
                    label="📥 Download DOCX",
                    data=st.session_state.generated_docx,
                    file_name=st.session_state.generated_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            with col2:
                if st.button("🔄 Regenerate", use_container_width=True):
                    # Clear and regenerate
                    st.session_state.pop("generated_docx", None)
                    st.session_state.pop("generated_content", None)
                    st.rerun()

            # Preview
            with st.expander("📄 Preview Generated Content"):
                st.markdown(st.session_state.generated_content)


if __name__ == "__main__":
    main()
