"""Configuration management for PDF to XLS converter."""

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


def get_api_key():
    """Get Anthropic API key from environment.

    Returns:
        str: The API key

    Raises:
        ValueError: If API key is not found or not set
    """
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key or api_key == 'your-api-key-here':
        raise ValueError(
            "ANTHROPIC_API_KEY not found or not set.\n"
            "Please edit the .env file and add your API key.\n"
            "Get your API key from: https://console.anthropic.com/"
        )
    return api_key


def get_model_name():
    """Get Claude model name from environment.

    Returns:
        str: The model name (defaults to claude-sonnet-4-5-20250929)
    """
    model = os.environ.get('CLAUDE_MODEL', 'claude-sonnet-4-5-20250929')
    return model
