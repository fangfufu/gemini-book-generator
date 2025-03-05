# gemini-book-generator
Use this to generate a book with Google Gemini

## Dependencies
The dependencies for this project are:
- ``google-generativeai``: The official Python client for interacting with the 
Gemini API.
- ``python-dotenv``: For loading environment variables from the .env file (for 
your API key).
- ``pypandoc``: A Python wrapper for the pandoc document converter. This makes 
it easier to call pandoc from our Python script.
- ``pandoc``: (Not a Python package!) This is the actual document converter. You 
must install this separately, as described in my previous response (it's not 
installed via pip).

## Use of .env file
You need to create a text file named ``.env`` with the following contents:
```
GOOGLE_API_KEY=your_actual_gemini_api_key
```
The ``.env`` file is in ``.gitignore``. It will not be uploaded to GitHub if you
fork the project.