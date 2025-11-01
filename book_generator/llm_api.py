import hashlib
import json
import logging
import os
import pathlib


def _check_for_repetition(
    current_text,
    repetition_check_config,
    full_response_text_parts_for_check=None,
):
    """
    Checks for repeated sentences or paragraphs in the generated text.

    Args:
        current_text (str): The current chunk of text from the LLM stream.
        repetition_check_config (dict): Configuration for repetition checks.
        full_response_text_parts_for_check (list, optional): A list containing
            the history of text parts for more robust checking. Defaults to None.

    Returns:
        bool: True if a repetition is detected, False otherwise.
    """
    min_sentence_length = repetition_check_config.get(
        "min_sentence_length_for_repetition_check", 10
    )
    max_history_size = repetition_check_config.get("max_history_size", 5)

    # Check 1: Exact repetition of the current chunk in the history
    if (
        full_response_text_parts_for_check is not None
        and len(current_text) > min_sentence_length
    ):
        # Count occurrences of the current chunk in the recent history
        recent_history = full_response_text_parts_for_check[-max_history_size:]
        if recent_history.count(current_text) > 1:
            logging.warning(
                f"Repetition detected (exact chunk): Found chunk '{current_text}' multiple times in recent history."
            )
            return True

    # Check 2: Repetition of sentences using the full text
    # This check is more expensive, so it's done on the accumulated text
    full_text = "".join(full_response_text_parts_for_check or [])
    if len(full_text) < min_sentence_length * 2:  # Not enough text to check
        return False

    # Simple sentence tokenization
    sentences = [
        s.strip() for s in full_text.replace("!", ".").replace("?", ".").split(".") if s
    ]
    long_sentences = [s for s in sentences if len(s) > min_sentence_length]

    if len(long_sentences) > len(set(long_sentences)):
        # Find the repeated sentence for logging purposes
        seen = set()
        for sentence in long_sentences:
            if sentence in seen:
                logging.warning(
                    f"Repetition detected (sentence): Sentence '{sentence}' is repeated."
                )
                return True
            seen.add(sentence)

    return False
import sys
import time

from google.generativeai.types import GenerationConfig
import google.generativeai as genai
import requests
from dotenv import load_dotenv
from transformers import AutoTokenizer

from book_generator.utils import sanitize_filename
from book_generator.constants import REPETITION_DETECTED


def setup_environment():
    """Loads environment variables from .env file and retrieves API key."""
    load_dotenv()
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        logging.error(
            "Error: GEMINI_API_KEY not found in .env file or environment variables."
        )
        sys.exit(1)
    logging.info("Environment variables loaded and API key found.")
    return api_key


# --- Caching Mechanism ---
def get_cache_path(prompt_text, cache_dir, cache_prefix=None):
    """Generates the cache file path, optionally prepending a prefix."""
    prompt_hash = hashlib.sha256(prompt_text.encode("utf-8")).hexdigest()
    pathlib.Path(cache_dir).mkdir(parents=True, exist_ok=True)

    filename_base = prompt_hash
    if cache_prefix:
        # Sanitize the prefix to make it filename-safe and limit length
        # Use a shorter length limit for prefixes to avoid overly long filenames
        sanitized_prefix = sanitize_filename(cache_prefix, max_length=50)
        if sanitized_prefix:  # Ensure sanitization didn't result in an empty string
            filename_base = f"{sanitized_prefix}_{prompt_hash}"
            logging.debug(f"Using cache prefix: '{sanitized_prefix}'")
        else:
            logging.error(
                f"Cache prefix '{cache_prefix}' sanitized to empty string. Using hash only."
            )

    return pathlib.Path(cache_dir) / f"{filename_base}.json"


def load_from_cache(prompt_text, cache_dir, cache_prefix=None):
    """Loads response from cache if available."""
    cache_file = get_cache_path(prompt_text, cache_dir, cache_prefix)
    if cache_file.exists():
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                cached_data = json.load(f)
            if "prompt" in cached_data and "response" in cached_data:
                # Log the actual filename for clarity
                logging.info(f"Cache hit for file: {cache_file.name}")
                return cached_data["response"]
            else:
                logging.error(f"Invalid cache file format: {cache_file}. Ignoring.")
                return None
        except Exception as e:
            logging.error(f"Error reading cache file {cache_file}: {e}")
            return None
    logging.debug(f"Cache miss for file: {cache_file.name}")
    return None


def save_to_cache(prompt_text, response_text, cache_dir, cache_prefix=None):
    """Saves the API response to the cache."""
    cache_file = get_cache_path(prompt_text, cache_dir, cache_prefix)
    try:
        cache_data = {"prompt": prompt_text, "response": response_text}
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=4)
        logging.info(f"Response saved to cache: {cache_file}")
    except Exception as e:
        logging.error(f"Error saving response to cache file {cache_file}: {e}")


# --- LLM API Interaction ---
def _call_gemini_api_internal(prompt, config, cache_prefix=None):
    """
    Internal function to call the Gemini API.
    Assumes caching is handled by the caller.

    Args:
        prompt (str): The prompt to send to the API.
        config (dict): The application configuration.
        cache_prefix (str, optional): A prefix to add to the cache filename
                                      for better identification. Defaults to None.
                                      (Note: cache_prefix is for logging/context here, actual caching is external)
    Returns:
        str or None: The API response text, or None if an error occurred.
    """
    api_settings_conf = config.get("api_settings", {})
    gemini_conf = api_settings_conf.get("gemini", {})

    default_max_retries = api_settings_conf.get("default_max_retries", 3)
    default_retry_delay = api_settings_conf.get("default_retry_delay_seconds", 5)

    # Logging for Gemini-specific call initiation (cache prefix is for context)
    model_name = gemini_conf.get("model", "gemini-2.0-flash-latest")
    temperature = float(
        gemini_conf.get("temperature", 1.0)
    )  # Default from example config.yaml

    max_retries = int(gemini_conf.get("max_retries", default_max_retries))
    retry_delay = int(gemini_conf.get("retry_delay_seconds", default_retry_delay))
    # safety_settings would be fetched from gemini_conf if specified in config.yaml under api_settings.gemini
    safety_settings = gemini_conf.get("safety_settings", None)

    verbose_debug = config.get("debug_options", {}).get("verbose_debug", False)
    stream_gemini = (
        verbose_debug  # Specifically for Gemini streaming if verbose_debug is on
    )

    try:
        if verbose_debug:
            logging.info(f"Gemini API Prompt for model '{model_name}':\n{prompt}")
            # For very long prompts, you might want to log only a portion or a summary
            # logging.info(f"Gemini API Prompt for model '{model_name}' (first 500 chars):\n{prompt[:500]}...")

        client = genai.Client()
        generation_config = GenerationConfig(
            temperature=temperature
        )

        # Count tokens for Gemini prompt
        try:
            token_count_response = client.models.count_tokens(
                model=model_name, contents=prompt
            )
            prompt_token_count = token_count_response.total_tokens
            logging.info(
                f"Gemini prompt token count for model '{model_name}': {prompt_token_count} tokens."
            )
        except Exception as e_token:
            logging.error(
                f"Could not count tokens for Gemini prompt (model '{model_name}'): {e_token}",
                exc_info=True
            )

        for attempt in range(max_retries):
            try:
                if stream_gemini:
                    response = client.models.generate_content_stream(
                        model=model_name,
                        contents=prompt,
                        config=generation_config,
                    )
                else:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt,
                        config=generation_config,
                    )

                if stream_gemini:
                    logging.info(f"Streaming Gemini response for model '{model_name}':")
                    full_response_text_parts = []
                    repetition_check_config = config.get("repetition_check", {})
                    print(f"\n--- Gemini Stream ({model_name}) ---")
                    for chunk in response:
                        if hasattr(chunk, "text"):
                            response_part = chunk.text
                            print(response_part, end="", flush=True)
                            full_response_text_parts.append(response_part)

                            if repetition_check_config.get("enabled", False):
                                if _check_for_repetition(
                                    response_part,
                                    repetition_check_config,
                                    full_response_text_parts,
                                ):
                                    logging.warning(
                                        "Repetition detected. Terminating and retrying."
                                    )
                                    print(
                                        "\n--- End Gemini Stream (Repetition Detected) ---"
                                    )
                                    return REPETITION_DETECTED
                        elif (
                            hasattr(chunk, "prompt_feedback")
                            and chunk.prompt_feedback
                            and chunk.prompt_feedback.block_reason
                        ):
                            logging.error(
                                f"Gemini API stream blocked. Reason: {chunk.prompt_feedback.block_reason}"
                            )
                            print(f"\n--- End Gemini Stream (Blocked) ---")
                            return None

                    print(f"\n--- End Gemini Stream (Done) ---")
                    logging.info(
                        f"Gemini API stream completed for model '{model_name}'."
                    )
                    if full_response_text_parts:
                        final_text = "".join(
                            str(p) for p in full_response_text_parts
                        ).strip()
                        return final_text
                    return ""
                else:  # Not streaming
                    response_text = None  # Initialize

                    # Handle prompt feedback first: if the prompt itself was blocked, don't retry.
                    if (
                        response.prompt_feedback
                        and response.prompt_feedback.block_reason
                    ):
                        logging.error(
                            f"Gemini API call blocked due to prompt. Reason: {response.prompt_feedback.block_reason}. Will not retry."
                        )
                        return None  # Explicitly do not retry prompt blocks

                    # If prompt was not blocked, check for response parts
                    if response.parts:
                        response_text = response.text
                    # Removed the elif for prompt_feedback.block_reason here as it's handled above.
                    # The case below is for when there are no parts, and it wasn't a prompt block.
                    # This could be due to finish_reason (e.g., SAFETY on response, MAX_TOKENS).
                    else:
                        # Check finish reason even if parts are empty
                        finish_reason = "UNKNOWN"
                        try:
                            # Access finish_reason safely
                            if response.candidates:
                                finish_reason = response.candidates[
                                    0
                                ].finish_reason.name  # Use .name for enum
                        except (AttributeError, IndexError):
                            logging.error(
                                "Could not determine finish reason from response."
                            )

                        logging.error(
                            f"API call returned no content or parts (and was not prompt-blocked). Finish Reason: {finish_reason}. Will attempt retry if applicable."
                        )
                        # response_text remains None

                    if response_text is not None:  # Check if we got valid text
                        logging.info(
                            f"Gemini API call successful for model {model_name}."
                        )
                        return response_text
                    else:
                        logging.error(
                            f"API attempt {attempt + 1} for model {model_name} resulted in no content (response_text is None). Will proceed to retry logic."
                        )

            except Exception as e:
                logging.error(
                    f"Gemini API call attempt {attempt + 1} for model {model_name} failed: {e}",
                    exc_info=True,
                )
                if "quota" in str(e).lower():  # Basic check for quota issues
                    logging.error(
                        f"Gemini API quota likely exceeded: {e}. Retrying as per configuration..."
                    )
                    # Removed 'return None' to allow retry for quota issues
                # For other general exceptions, the loop will continue to the retry logic.

            if attempt < max_retries - 1:
                logging.info(f"Retrying Gemini API call in {retry_delay} seconds...")
                time.sleep(retry_delay)

        logging.error(
            f"Gemini API call for model {model_name} failed after {max_retries} attempts."
        )
        if api_settings_conf.get("terminate_on_api_failure", False):
            logging.critical("Terminating program due to repeated API failures.")
            sys.exit(1)
        return None  # Explicitly return None after all retries fail

    except Exception as e:
        logging.error(f"An unexpected error occurred during Gemini API call setup: {e}")
        if api_settings_conf.get("terminate_on_api_failure", False):
            logging.critical("Terminating program due to unexpected API error.")
            sys.exit(1)
        return None


def _call_ollama_api_internal(prompt, config, cache_prefix=None):
    """
    Internal function to call the Ollama API.
    Assumes caching is handled by the caller.

    Args:
        prompt (str): The prompt to send to the API.
        config (dict): The application configuration.
        cache_prefix (str, optional): Contextual prefix, caching is external.

    Returns:
        str or None: The API response text, or None if an error occurred.
    """
    api_settings_conf = config.get("api_settings", {})
    ollama_config = api_settings_conf.get("ollama", {})

    default_max_retries = api_settings_conf.get("default_max_retries", 3)
    default_retry_delay = api_settings_conf.get("default_retry_delay_seconds", 5)

    base_url = ollama_config.get("base_url", "http://localhost:11434")
    model_name = ollama_config.get("model", "llama3")  # Default Ollama model
    tokenizer_model_name = ollama_config.get(
        "tokenizer_model", "NousResearch/Llama-3-8B-Instruct-hf"
    )  # Default Llama3 tokenizer

    max_retries = int(ollama_config.get("max_retries", default_max_retries))
    retry_delay = int(ollama_config.get("retry_delay_seconds", default_retry_delay))
    request_timeout = int(
        ollama_config.get("request_timeout_seconds", 120)
    )  # Default 2 mins
    api_url = f"{base_url.rstrip('/')}/api/generate"

    verbose_debug = config.get("debug_options", {}).get("verbose_debug", False)
    stream_ollama = (
        verbose_debug  # Specifically for Ollama streaming if verbose_debug is on
    )

    # Prepare payload, starting with basic info
    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": stream_ollama,
        "options": {},  # Initialize empty options
    }

    # Get all llm_options from the config
    llm_options = ollama_config.get("llm_options", {})
    if llm_options:
        logging.info(f"Applying Ollama llm_options: {llm_options}")
        # Iterate through the provided llm_options and add them to the payload's options
        for key, value in llm_options.items():
            if value is not None:  # Ensure not to add keys with None value
                payload["options"][key] = value
                logging.debug(f"Set Ollama option '{key}': {value}")

    # For backward compatibility, check for standalone temperature if not in llm_options
    if "temperature" not in payload["options"] and "temperature" in ollama_config:
        payload["options"]["temperature"] = float(ollama_config["temperature"])
        logging.info(
            f"Applying standalone 'temperature' setting: {payload['options']['temperature']}"
        )

    # For backward compatibility, handle standalone context_window_size
    if "num_ctx" not in payload["options"] and "context_window_size" in ollama_config:
        context_window_size = ollama_config.get("context_window_size")
        if context_window_size is not None:
            try:
                payload["options"]["num_ctx"] = int(context_window_size)
                logging.info(
                    f"Applying standalone 'context_window_size' as num_ctx: {payload['options']['num_ctx']}"
                )
            except (ValueError, TypeError):
                logging.error(
                    f"Invalid 'context_window_size' value: {context_window_size}. It must be an integer. Ignoring."
                )

    if verbose_debug:
        logging.info(f"Ollama API Prompt for model '{model_name}':\n{prompt}")
        # For very long prompts, you might want to log only a portion or a summary
        # logging.info(f"Ollama API Prompt for model '{model_name}' (first 500 chars):\n{prompt[:500]}...")

    # Attempt client-side token counting for Ollama
    # Note: For higher efficiency with many calls, consider loading the tokenizer once outside this function.
    if tokenizer_model_name:
        try:
            logging.debug(
                f"Loading tokenizer: {tokenizer_model_name} for Ollama prompt token count."
            )
            hugging_face_token = os.getenv("HF_TOKEN")
            tokenizer = AutoTokenizer.from_pretrained(
                tokenizer_model_name, token=hugging_face_token
            )
            token_ids = tokenizer.encode(prompt)
            num_tokens = len(token_ids)
            logging.info(
                f"Ollama client-side token count for prompt (tokenizer: '{tokenizer_model_name}', model: '{model_name}'): {num_tokens} tokens."
            )
        except Exception as e_token_ollama:
            logging.error(
                f"Could not count tokens for Ollama prompt using tokenizer '{tokenizer_model_name}' (model: '{model_name}'): {e_token_ollama}",
                exc_info=True,
            )
            logging.info(
                f"Ollama API for model '{model_name}': Standard Ollama API does not provide a direct prompt token count. Client-side estimation failed."
            )
    else:
        logging.info(
            f"Ollama API for model '{model_name}': No tokenizer_model configured for client-side token counting. Standard Ollama API does not provide a direct prompt token count."
        )

    for attempt in range(max_retries):
        try:
            response = requests.post(
                api_url,
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=request_timeout,
                stream=stream_ollama,  # Pass stream=True to requests.post if streaming
            )
            response.raise_for_status()  # Raises HTTPError for bad responses (4XX, 5XX)

            if stream_ollama:
                logging.info(f"Streaming Ollama response for model '{model_name}':")
                full_response_text_parts = []
                repetition_check_config = config.get("repetition_check", {})
                print(f"\n--- Ollama Stream ({model_name}) ---")
                for line in response.iter_lines():
                    if line:
                        decoded_line = line.decode("utf-8")
                        try:
                            chunk = json.loads(decoded_line)
                            if "error" in chunk:
                                logging.error(
                                    f"Ollama API error during stream for model '{model_name}': {chunk['error']}"
                                )
                                print(f"\n--- End Ollama Stream (Error) ---")
                                return None

                            response_part = chunk.get("response", "")
                            print(response_part, end="", flush=True)
                            full_response_text_parts.append(response_part)

                            # Perform repetition check if enabled
                            if repetition_check_config.get("enabled", False):
                                if _check_for_repetition(
                                    response_part,
                                    repetition_check_config,
                                    full_response_text_parts,
                                ):
                                    logging.warning(
                                        "Repetition detected. Terminating and retrying."
                                    )
                                    print(
                                        "\n--- End Ollama Stream (Repetition Detected) ---"
                                    )
                                    return REPETITION_DETECTED

                            if chunk.get("done"):
                                print(f"\n--- End Ollama Stream (Done) ---")
                                logging.info(
                                    f"Ollama API stream completed for model '{model_name}'."
                                )
                                final_text = "".join(full_response_text_parts).strip()
                                return final_text
                        except json.JSONDecodeError:
                            logging.error(
                                f"Error decoding JSON chunk from Ollama stream: {decoded_line}"
                            )
                            print(f"\n--- End Ollama Stream (JSON Error) ---")
                            return None
                print(f"\n--- End Ollama Stream (Unexpected End) ---")
                logging.warning(
                    "Ollama stream ended without a 'done: true' message."
                )
                return (
                    "".join(full_response_text_parts).strip()
                    if full_response_text_parts
                    else None
                )
            else:  # Not streaming
                response_data = response.json()
                if "error" in response_data:
                    logging.error(
                        f"Ollama API error for model '{model_name}': {response_data['error']}"
                    )
                    return None
                if "response" in response_data:
                    response_text = response_data["response"]
                    logging.info(
                        f"Ollama API call successful for model '{model_name}'."
                    )
                    return response_text.strip()
                else:
                    logging.error(
                        f"Ollama API response for model '{model_name}' did not contain 'response' key. Attempt {attempt + 1}/{max_retries}. Data: {response_data}"
                    )

        except requests.exceptions.HTTPError as e:
            logging.error(
                f"Ollama API call (model '{model_name}') attempt {attempt + 1} failed with HTTPError: {e}. Status: {e.response.status_code}",
                exc_info=True,
            )
            if e.response.status_code == 404:  # Model not found
                try:
                    error_detail = e.response.json().get("error", "Model not found")
                    logging.error(
                        f"Ollama model '{model_name}' not found: {error_detail}. Please ensure the model is pulled and available."
                    )
                except json.JSONDecodeError:
                    logging.error(f"Ollama model '{model_name}' not found (404).")
                return None  # Don't retry if model not found
        except (
            requests.exceptions.RequestException
        ) as e:  # Covers ConnectionError, Timeout, etc.
            logging.error(
                f"Ollama API call (model '{model_name}') attempt {attempt + 1} failed: {e}",
                exc_info=True,
            )

        if attempt < max_retries - 1:
            logging.info(f"Retrying Ollama API call in {retry_delay} seconds...")
            time.sleep(retry_delay)
        else:
            logging.error(
                f"Ollama API call for model '{model_name}' failed after {max_retries} attempts."
            )
            if api_settings_conf.get("terminate_on_api_failure", False):
                logging.critical(
                    "Terminating program due to repeated API failures."
                )
                sys.exit(1)
            return None
    return None  # Should be covered by loop logic, but as a safeguard


def call_llm_api(prompt, config, cache_prefix=None):
    """
    Calls the configured LLM API (Gemini or Ollama), using caching and handling retries.
    Includes special handling for repetition detection.
    """
    cache_dir = config.get("cache_dir", "api_cache")
    cached_response = load_from_cache(prompt, cache_dir, cache_prefix)
    if cached_response is not None:
        return cached_response

    api_settings = config.get("api_settings", {})
    api_provider = api_settings.get("provider", "gemini")
    logging.info(
        f"Calling {api_provider.upper()} API... (Cache Prefix: {cache_prefix or 'None'})"
    )

    repetition_check_config = config.get("repetition_check", {})
    max_retries_repetition = repetition_check_config.get(
        "max_retries_on_repetition", 2
    )
    repetition_retry_delay = repetition_check_config.get(
        "repetition_retry_delay_seconds", 5
    )
    response_text = None

    for attempt in range(max_retries_repetition + 1):
        if api_provider == "gemini":
            response_text = _call_gemini_api_internal(prompt, config, cache_prefix)
        elif api_provider == "ollama":
            response_text = _call_ollama_api_internal(prompt, config, cache_prefix)
        else:
            logging.error(f"Unsupported API provider: {api_provider}")
            return None

        if response_text == REPETITION_DETECTED:
            logging.warning(
                f"Repetition detected in response. Attempt {attempt + 1}/{max_retries_repetition + 1}."
            )
            if attempt < max_retries_repetition:
                logging.info(
                    f"Retrying after {repetition_retry_delay} seconds due to repetition..."
                )
                time.sleep(repetition_retry_delay)
                # Invalidate cache for this specific prompt before retrying
                cache_file = get_cache_path(prompt, cache_dir, cache_prefix)
                if cache_file.exists():
                    os.remove(cache_file)
            else:
                logging.error(
                    "Max retries for repetition reached. Returning None."
                )
                return None
        else:
            # If not a repetition or another error, break the loop
            break

    if response_text is not None and response_text != REPETITION_DETECTED:
        save_to_cache(prompt, response_text, cache_dir, cache_prefix)

    return response_text
