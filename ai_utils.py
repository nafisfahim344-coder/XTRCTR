import os
import google.generativeai as genai
from pathlib import Path
import time

class XTRCTRAI:
    def __init__(self, api_key=None):
        self.api_key = api_key
        if api_key:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel("gemini-1.5-flash")
        else:
            self.model = None

    def is_configured(self):
        return self.api_key is not None

    def analyze_pdf_structure(self, pdf_path):
        """
        Uploads PDF to Gemini and asks for section detection.
        Returns a list of dicts: [{'title': ..., 'start': ..., 'end': ...}]
        """
        if not self.model:
            return None, "API Key not set"

        try:
            # Upload file
            sample_file = genai.upload_file(path=pdf_path, display_name="PDF for Sectioning")
            
            # Wait for processing
            while sample_file.state.name == "PROCESSING":
                time.sleep(2)
                sample_file = genai.get_file(sample_file.name)

            if sample_file.state.name == "FAILED":
                return None, "File processing failed on Gemini servers"

            prompt = """
            Analyze this PDF and identify its logical sections (Chapters, Sections, or Parts).
            For each section, provide:
            1. A concise, professional title (for use as a filename).
            2. The starting page number.
            3. The ending page number.

            Respond ONLY with a JSON list of objects like this:
            [{"title": "Introduction", "start": 1, "end": 5}, {"title": "Chapter 1", "start": 6, "end": 20}]
            """

            response = self.model.generate_content([sample_file, prompt])
            
            # Simple JSON extraction (Gemini often wraps in ```json)
            text = response.text
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()
            
            import json
            sections = json.loads(text)
            
            # Clean up file
            genai.delete_file(sample_file.name)
            
            return sections, "success"

        except Exception as e:
            return None, str(e)

    def suggest_filename(self, pdf_path, start_page, end_page):
        """
        Suggests a professional filename for a specific page range.
        """
        if not self.model:
            return None

        try:
            # For speed, we might just send the first page of the range as an image
            # but for a prototype, let's just use text if possible, or Gemini can handle the PDF range if we extract it.
            # However, simplest prototype is to trust the structure analysis.
            pass
        except Exception:
            return None
