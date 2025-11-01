import unittest
from unittest.mock import patch, MagicMock
from book_generator.llm_api import call_llm_api
from book_generator.constants import REPETITION_DETECTED

class TestLlmApi(unittest.TestCase):
    @patch("os.remove")
    @patch("book_generator.llm_api.get_cache_path")
    @patch("book_generator.llm_api.load_from_cache")
    @patch("book_generator.llm_api._call_ollama_api_internal")
    def test_repetition_detection_and_retry_ollama(
        self,
        mock_call_ollama,
        mock_load_from_cache,
        mock_get_cache_path,
        mock_os_remove,
    ):
        # Arrange
        mock_load_from_cache.return_value = None
        mock_call_ollama.side_effect = [REPETITION_DETECTED, "successful response"]
        mock_get_cache_path.return_value = MagicMock()
        mock_get_cache_path.return_value.exists.return_value = True

        config = {
            "api_settings": {"provider": "ollama"},
            "repetition_check": {
                "enabled": True,
                "max_retries_on_repetition": 1,
                "repetition_retry_delay_seconds": 0,
            },
        }
        prompt = "test prompt"

        # Act
        result = call_llm_api(prompt, config)

        # Assert
        self.assertEqual(result, "successful response")
        self.assertEqual(mock_call_ollama.call_count, 2)
        mock_os_remove.assert_called_once_with(mock_get_cache_path.return_value)

    @patch("os.remove")
    @patch("book_generator.llm_api.get_cache_path")
    @patch("book_generator.llm_api.load_from_cache")
    @patch("book_generator.llm_api._call_gemini_api_internal")
    def test_repetition_detection_and_retry_gemini(
        self,
        mock_call_gemini,
        mock_load_from_cache,
        mock_get_cache_path,
        mock_os_remove,
    ):
        # Arrange
        mock_load_from_cache.return_value = None
        mock_call_gemini.side_effect = [REPETITION_DETECTED, "successful response"]
        mock_get_cache_path.return_value = MagicMock()
        mock_get_cache_path.return_value.exists.return_value = True

        config = {
            "api_settings": {"provider": "gemini"},
            "repetition_check": {
                "enabled": True,
                "max_retries_on_repetition": 1,
                "repetition_retry_delay_seconds": 0,
            },
        }
        prompt = "test prompt"

        # Act
        result = call_llm_api(prompt, config)

        # Assert
        self.assertEqual(result, "successful response")
        self.assertEqual(mock_call_gemini.call_count, 2)
        mock_os_remove.assert_called_once_with(mock_get_cache_path.return_value)

if __name__ == "__main__":
    unittest.main()
