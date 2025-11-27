"""Mermaid diagram renderer using mermaid-cli."""

import asyncio
import logging
from pathlib import Path
from typing import Optional

try:
    from mermaid_cli import render_mermaid
except ImportError:
    render_mermaid = None

logger = logging.getLogger(__name__)


class MermaidRenderer:
    """Render Mermaid diagrams to PNG images using mermaid-cli."""

    def __init__(self):
        """Initialize the Mermaid renderer."""
        if render_mermaid is None:
            raise ImportError(
                "mermaid-cli is not installed. "
                "Install it with: pip install mermaid-cli && playwright install chromium"
            )

    def render_to_file(self, mermaid_code: str, output_path: str) -> bool:
        """
        Render Mermaid diagram to a PNG file.

        Args:
            mermaid_code: Mermaid diagram definition
            output_path: Path to save the PNG file

        Returns:
            True if rendering was successful, False otherwise
        """
        try:
            # Run the async rendering function
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                result = loop.run_until_complete(
                    self._render_async(mermaid_code, output_path)
                )
                return result
            finally:
                loop.close()
        except Exception as e:
            logger.error(f"Failed to render Mermaid diagram: {e}")
            return False

    async def _render_async(self, mermaid_code: str, output_path: str) -> bool:
        """
        Async helper to render Mermaid diagram.

        Args:
            mermaid_code: Mermaid diagram definition
            output_path: Path to save the PNG file

        Returns:
            True if rendering was successful, False otherwise
        """
        try:
            # Configure Mermaid with Japanese font support
            mermaid_config = {
                "themeVariables": {
                    "fontFamily": '"Noto Sans JP", "Hiragino Sans", "Yu Gothic", "Meiryo", sans-serif'
                }
            }

            # Render the diagram
            title, description, diagram_bytes = await render_mermaid(
                definition=mermaid_code,
                output_format="png",
                background_color="white",
                mermaid_config=mermaid_config
            )

            # Save to file
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            output_file.write_bytes(diagram_bytes)

            logger.info(f"Successfully rendered Mermaid diagram to {output_path}")
            return True

        except Exception as e:
            logger.error(f"Error rendering Mermaid diagram: {e}")
            return False
