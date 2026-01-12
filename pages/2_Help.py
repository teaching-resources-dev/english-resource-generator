"""
Help Page

Usage instructions for the English Resource Generator.
"""

import streamlit as st

st.set_page_config(
    page_title="Help",
    page_icon="ℹ️",
    layout="centered"
)

st.title("ℹ️ Help")

st.markdown("""
## How to Use This App

### 1. Login
Enter your authorised school email address on the home page.

### 2. Generate Resources
1. Click **Generate** in the sidebar
2. Select the resource type (Quiz, Worked Example, etc.)
3. Choose your year level and text
4. Enter the topic or focus area
5. Click **Generate**
6. Download the Word document

### Available Resource Types

| Resource | Description |
|----------|-------------|
| **Quiz** | Multiple-choice questions with answer key |
| *Coming soon* | Worked examples, Grammar worksheets, 4-Step analysis |

### Tips for Best Results

- **Be specific with topics**: "Character development of Nathaniel" works better than "characters"
- **Check the preview**: Review the generated content before downloading
- **Regenerate if needed**: Use the Regenerate button if the output isn't quite right

### Need Help?

Contact your administrator if you:
- Can't log in (email not authorised)
- Experience repeated errors
- Need a new text added to the system

---

*English Resource Generator v0.1*
""")
