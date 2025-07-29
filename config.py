"""
Configuration module for Formatly V3.
Handles secure API key management and application settings.
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from typing import Optional, Dict, Any

# Load environment variables
load_dotenv()

class Config:
    """Configuration class for Formatly V3."""
    
    def __init__(self):
        print(f"[DEBUG] Config.__init__: Initializing configuration...")
        self.api_key = self._get_api_key()
        print(f"[DEBUG] Config.__init__: API key configured: {'Yes' if self.api_key else 'No'}")
        self.model_name = os.getenv("GEMINI_MODEL", "gemini-2.0-flash-exp")
        print(f"[DEBUG] Config.__init__: Model name: {self.model_name}")
        self.max_retries = int(os.getenv("MAX_RETRIES", "3"))
        self.timeout = int(os.getenv("TIMEOUT", "30"))
        self.log_level = os.getenv("LOG_LEVEL", "INFO")
        print(f"[DEBUG] Config.__init__: Max retries: {self.max_retries}, Timeout: {self.timeout}s")
        
        # Gemini API settings
        self.temperature = float(os.getenv("GEMINI_TEMPERATURE", "1"))
        self.top_p = float(os.getenv("GEMINI_TOP_P", "0.95"))
        self.top_k = int(os.getenv("GEMINI_TOP_K", "40"))
        self.max_output_tokens = int(os.getenv("GEMINI_MAX_OUTPUT_TOKENS", "1024"))
        self.stop_sequences = os.getenv("GEMINI_STOP_SEQUENCES", "\n\n").split(',')
        
    def _get_api_key(self) -> Optional[str]:
        """Securely retrieve API key from environment variables."""
        api_key = os.getenv("GEMINI_API_KEY")
        print(f"[DEBUG] Config._get_api_key: Retrieved API key length: {len(api_key) if api_key else 0}")
        return api_key
    
    def is_api_key_configured(self) -> bool:
        """Check if API key is configured."""
        configured = self.api_key is not None and len(self.api_key.strip()) > 0
        print(f"[DEBUG] Config.is_api_key_configured: {configured}")
        return configured
    
    def get_safe_config(self) -> Dict[str, Any]:
        """Get configuration without sensitive data for logging."""
        safe_config = {
            "model_name": self.model_name,
            "max_retries": self.max_retries,
            "timeout": self.timeout,
            "log_level": self.log_level,
            "api_key_configured": self.is_api_key_configured()
        }
        print(f"[DEBUG] Config.get_safe_config: {safe_config}")
        return safe_config

# Global configuration instance
config = Config()

# Supported file formats
SUPPORTED_FORMATS = [".docx"]

# Supported citation styles
SUPPORTED_STYLES = ["apa", "mla", "chicago"]

# Default settings
DEFAULT_STYLE = "apa"
DEFAULT_OUTPUT_SUFFIX = "formatted"

# Application metadata
APP_NAME = "Formatly V3"
APP_VERSION = "3.0.0"
APP_DESCRIPTION = "AI-powered academic document formatter"
