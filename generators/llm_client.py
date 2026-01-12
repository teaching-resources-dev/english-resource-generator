"""
LLM Client for English Resource Generator.

Wraps Google Gemini API with error handling and retry logic.
"""

import streamlit as st
import google.generativeai as genai
from typing import Optional


class GeminiClient:
    """Client for Google Gemini API."""

    def __init__(self):
        """Initialise the Gemini client with API key from secrets."""
        try:
            api_key = st.secrets.gemini.api_key
        except (KeyError, AttributeError):
            # For local development
            import os
            api_key = os.getenv("GEMINI_API_KEY")
            if not api_key:
                raise ValueError(
                    "GEMINI_API_KEY not found. Set it in .streamlit/secrets.toml "
                    "or as an environment variable."
                )

        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel("gemini-1.5-flash")

    def generate(
        self,
        prompt: str,
        max_tokens: int = 4000,
        temperature: float = 0.7
    ) -> str:
        """
        Generate content using Gemini.

        Args:
            prompt: The prompt to send to the model.
            max_tokens: Maximum tokens in response.
            temperature: Creativity level (0.0-1.0).

        Returns:
            Generated text content.

        Raises:
            Exception: If generation fails after retries.
        """
        try:
            response = self.model.generate_content(
                prompt,
                generation_config={
                    "max_output_tokens": max_tokens,
                    "temperature": temperature
                }
            )
            return response.text
        except Exception as e:
            error_msg = str(e).lower()
            if "rate" in error_msg or "quota" in error_msg:
                raise RateLimitError("Rate limit reached. Please wait a moment and try again.")
            elif "safety" in error_msg or "blocked" in error_msg:
                raise ContentFilterError("Content was filtered. Try rephrasing the topic.")
            else:
                raise GenerationError(f"Generation failed: {str(e)}")


class RateLimitError(Exception):
    """Raised when API rate limit is hit."""
    pass


class ContentFilterError(Exception):
    """Raised when content is filtered by safety settings."""
    pass


class GenerationError(Exception):
    """Raised when generation fails for other reasons."""
    pass


@st.cache_resource
def get_llm_client() -> GeminiClient:
    """Get a cached LLM client instance."""
    return GeminiClient()
