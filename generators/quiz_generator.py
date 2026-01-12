"""
Quiz Generator for English Resource Generator.

Generates multiple-choice quizzes using Gemini and converts to DOCX.
"""

import json
import re
from pathlib import Path
from io import BytesIO
from typing import Dict, Optional

import streamlit as st

from generators.llm_client import get_llm_client, RateLimitError, ContentFilterError, GenerationError
from docx_generation.docx_styles import (
    setup_document,
    add_header_footer,
    add_title,
    add_subtitle,
    add_subsection_heading,
    add_question,
    add_body_paragraph,
    add_content_table,
    add_horizontal_rule,
    add_page_break,
    save_document
)


def load_prompt_template() -> str:
    """Load the quiz prompt template."""
    prompt_path = Path(__file__).parent.parent / "prompts" / "quiz.txt"
    return prompt_path.read_text(encoding="utf-8")


def load_pedagogy_core() -> str:
    """Load the core pedagogy requirements."""
    pedagogy_path = Path(__file__).parent.parent / "prompts" / "pedagogy_core.txt"
    return pedagogy_path.read_text(encoding="utf-8")


def load_text_index() -> Dict:
    """Load the text index."""
    index_path = Path(__file__).parent.parent / "knowledge" / "index.json"
    with open(index_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_text_info(text_id: str) -> Optional[Dict]:
    """Get information about a specific text."""
    index = load_text_index()
    for text in index["texts"]:
        if text["id"] == text_id:
            return text
    return None


def generate_quiz(
    year_level: str,
    text_id: str,
    topic: str,
    num_questions: int = 10
) -> str:
    """
    Generate a quiz using Gemini.

    Args:
        year_level: Year level (e.g., "10", "11")
        text_id: ID of the text from index.json
        topic: Quiz topic
        num_questions: Number of questions (5-15)

    Returns:
        Generated quiz content as markdown.
    """
    # Get text info
    text_info = get_text_info(text_id)
    if not text_info:
        raise ValueError(f"Text not found: {text_id}")

    # Load templates
    prompt_template = load_prompt_template()
    pedagogy_core = load_pedagogy_core()

    # Build knowledge context (placeholder - will be enhanced later)
    knowledge_context = f"""
Text: {text_info['name']}
Type: {text_info['type']}
Year Level: Year {year_level}
Available Resources: {', '.join(text_info.get('resources', []))}
"""

    # Format the prompt
    prompt = prompt_template.format(
        year_level=f"Year {year_level}",
        text_name=text_info["name"],
        text_type=text_info["type"],
        topic=topic,
        num_questions=num_questions,
        pedagogy_core=pedagogy_core,
        knowledge_context=knowledge_context
    )

    # Generate with LLM
    client = get_llm_client()
    return client.generate(prompt, max_tokens=4000, temperature=0.7)


def parse_quiz_content(content: str) -> Dict:
    """
    Parse generated quiz content into structured data.

    Args:
        content: Raw markdown content from LLM.

    Returns:
        Dictionary with questions, answers, and metadata.
    """
    result = {
        "title": "",
        "text_name": "",
        "year_level": "",
        "questions": [],
        "answer_key": [],
        "teacher_notes": ""
    }

    # Extract title
    title_match = re.search(r"# Quiz: (.+)", content)
    if title_match:
        result["title"] = title_match.group(1).strip()

    # Extract text and year level
    text_match = re.search(r"\*\*Text\*\*: (.+)", content)
    if text_match:
        result["text_name"] = text_match.group(1).strip()

    year_match = re.search(r"\*\*Year Level\*\*: (.+)", content)
    if year_match:
        result["year_level"] = year_match.group(1).strip()

    # Extract questions
    question_pattern = r"### Question (\d+)\n(.+?)\n\nA\) (.+?)\nB\) (.+?)\nC\) (.+?)\nD\) (.+?)(?=\n---|\n### |$)"
    questions = re.findall(question_pattern, content, re.DOTALL)

    for q in questions:
        result["questions"].append({
            "number": int(q[0]),
            "text": q[1].strip(),
            "options": {
                "A": q[2].strip(),
                "B": q[3].strip(),
                "C": q[4].strip(),
                "D": q[5].strip()
            }
        })

    # Extract answer key
    answer_pattern = r"\| (\d+) \| ([A-D]) \| (.+?) \|"
    answers = re.findall(answer_pattern, content)

    for a in answers:
        result["answer_key"].append({
            "question": int(a[0]),
            "answer": a[1],
            "explanation": a[2].strip()
        })

    # Extract teacher notes (everything after "## Teacher Notes")
    notes_match = re.search(r"## Teacher Notes\n(.+)", content, re.DOTALL)
    if notes_match:
        result["teacher_notes"] = notes_match.group(1).strip()

    return result


def create_quiz_docx(quiz_data: Dict, year_level: str, text_name: str) -> BytesIO:
    """
    Create a DOCX file from parsed quiz data.

    Args:
        quiz_data: Parsed quiz dictionary.
        year_level: Year level for header.
        text_name: Text name for header.

    Returns:
        BytesIO object containing the DOCX file.
    """
    # Create document
    doc = setup_document(colour_scheme='professional_minimal')

    # Add header/footer
    add_header_footer(
        doc,
        year_level=year_level,
        unit_name=text_name,
        doc_type="Quiz",
        include_name=True
    )

    # Add title
    add_title(doc, f"Quiz: {quiz_data.get('title', 'Assessment')}")
    add_subtitle(doc, f"{text_name}")

    # Add questions
    add_subsection_heading(doc, "Questions")

    for q in quiz_data.get("questions", []):
        # Add question
        add_question(doc, q["number"], q["text"])

        # Add options
        for letter, option_text in q.get("options", {}).items():
            para = doc.add_paragraph()
            para.paragraph_format.left_indent = 635000  # ~0.5 inch
            run = para.add_run(f"{letter}) {option_text}")
            run.font.name = "Aptos"
            run.font.size = 132000  # 11pt in EMUs

        # Add spacing
        doc.add_paragraph()

    # Add page break before answer key
    add_page_break(doc)

    # Add answer key
    add_subsection_heading(doc, "Answer Key")

    # Create answer table
    headers = ["Q", "Answer", "Explanation"]
    rows = []
    for a in quiz_data.get("answer_key", []):
        rows.append([
            str(a["question"]),
            a["answer"],
            a["explanation"]
        ])

    if rows:
        add_content_table(doc, headers, rows)

    # Add teacher notes if present
    if quiz_data.get("teacher_notes"):
        add_horizontal_rule(doc)
        add_subsection_heading(doc, "Teacher Notes")
        add_body_paragraph(doc, quiz_data["teacher_notes"])

    # Save to BytesIO
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


def generate_quiz_docx(
    year_level: str,
    text_id: str,
    topic: str,
    num_questions: int = 10
) -> tuple[BytesIO, str]:
    """
    Generate a complete quiz DOCX file.

    Args:
        year_level: Year level.
        text_id: Text ID from index.
        topic: Quiz topic.
        num_questions: Number of questions.

    Returns:
        Tuple of (BytesIO docx file, raw markdown content).
    """
    # Get text info
    text_info = get_text_info(text_id)
    if not text_info:
        raise ValueError(f"Text not found: {text_id}")

    # Generate quiz content
    raw_content = generate_quiz(year_level, text_id, topic, num_questions)

    # Parse content
    quiz_data = parse_quiz_content(raw_content)

    # If parsing failed to get questions, use defaults
    if not quiz_data.get("title"):
        quiz_data["title"] = topic
    if not quiz_data.get("questions"):
        # Return raw content for manual review
        raise GenerationError(
            "Could not parse quiz structure. Please try regenerating or "
            "check the preview for issues."
        )

    # Create DOCX
    docx_buffer = create_quiz_docx(
        quiz_data,
        year_level=f"Year {year_level}",
        text_name=text_info["name"]
    )

    return docx_buffer, raw_content
