import os
import google.generativeai as genai
from pathlib import Path
import time
import json

class XTRCTRAI:
    def __init__(self, api_key=None):
        self.api_key = api_key
        if api_key:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel("gemini-2.5-flash")
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
            sample_file = genai.upload_file(path=pdf_path, display_name="PDF for Sectioning")
            
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
            
            text = response.text
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()
            
            sections = json.loads(text)
            genai.delete_file(sample_file.name)
            
            return sections, "success"

        except Exception as e:
            return None, str(e)

    def classify_document(self, ocr_text, target_program="Auto"):
        """
        Takes raw OCR text containing MULTIPLE PAGES and uses Gemini to classify 
        and SPLIT the document based on strict rules for Bangladeshi Korean 
        University applications.
        """
        if not self.model:
            return None, "API Key not set"

        strict_instruction = ""
        if target_program == "EAP/KLP PROGRAMME (Set 1)":
            strict_instruction = """
🔥🔥🔥 CRITICAL INSTRUCTION 🔥🔥🔥
The user has EXPLICITLY stated that all documents in this PDF belong to the EAP/KLP PROGRAMME (Set 1).
You MUST ONLY use the PDF Naming Convention from SET 1.
Do NOT attempt to guess the program. Do NOT use any names from Set 2.
Assign the program field as exactly "EAP/KLP".
"""
        elif target_program == "BACHELOR'S / MASTER'S (Set 2)":
            strict_instruction = """
🔥🔥🔥 CRITICAL INSTRUCTION 🔥🔥🔥
The user has EXPLICITLY stated that all documents in this PDF belong to the BACHELOR'S / MASTER'S PROGRAMME (Set 2).
You MUST ONLY use the PDF Naming Convention from SET 2.
Do NOT attempt to guess the program. Do NOT use any names from Set 1.
Assign the program field as exactly "BACHELOR'S/MASTER'S".
"""

        system_prompt = f"""
You are a document classification and splitting AI for Korean university 
applications from Bangladesh.

You will receive OCR-extracted text from a merged PDF file containing 
MULTIPLE documents. The text is separated by page markers like "--- PAGE 1 ---".

{strict_instruction}

Your job:
1. Identify each distinct document within the text.
2. Determine the exact start_page and end_page for each document based on the rules.
3. Identify WHICH PROGRAM the document belongs to.
4. Assign the correct official PDF name from that program's list.

════════════════════════════════════════
PROGRAM IDENTIFICATION RULES
════════════════════════════════════════
You must scan the ENTIRE document text.
- RULE 1: If there is ANY mention of a "Statement of Purpose", "SOP", or "Recommendation Letter" anywhere in the entire document, the PROGRAM IS DEFINITELY `BACHELOR'S/MASTER'S`.
- RULE 2: If there is NO SOP and NO Recommendation Letter in the entire document, the PROGRAM IS DEFINITELY `EAP/KLP`.

════════════════════════════════════════
SET 1 — EAP/KLP PROGRAMME
PDF Naming Convention & MERGE RULES:
════════════════════════════════════════
For EAP/KLP, you must NOT output individual documents (like NID, Birth Certificate separately). 
You must group them into these EXACT 5 MERGED PDF NAMES:

1.  PASSPORT
    -> Start page to end page of the passport.
2.  FAMILY RELATIONSHIP
    -> This covers Affidavit, Student NID, Student Birth Cert, Father NID, Mother NID, and Family Relationship Cert.
    -> The start_page should be the very first page of the family docs, and the end_page should be the last page of the family docs.
    -> Output ONLY ONE JSON object named "FAMILY RELATIONSHIP".
3.  ACADEMIC DOCUMENTS
    -> Covers all transcripts and certificates.
    -> Output ONLY ONE JSON object named "ACADEMIC DOCUMENTS".
4.  FINANCIAL DOCUMENTS
    -> Covers Bank Statement, Solvency, TIN, TAX, Trade License, Employment Cert.
    -> The start_page should be the bank statement, and the end_page should be the last tax document.
    -> Output ONLY ONE JSON object named "FINANCIAL DOCUMENTS".
5.  LANGUAGE PROFICIENCY
    -> Only if visible (IELTS, TOPIK, TOEFL).

════════════════════════════════════════
SET 2 — BACHELOR'S / MASTER'S PROGRAMME
PDF Naming Convention (DO NOT MERGE):
════════════════════════════════════════
For Bachelor's/Master's, output these exact individual documents separately:

1.  PASSPORT (COLOUR PRINT + NOTARIZED)
2.  STUDENT NID (TRANSLATION, NOTARY)
3.  STUDENT BIRTH CERTIFICATE (PHOTO COPY NOTARY)
4.  FATHER NID/DEATH CERTIFICATE (TRANSLATION, NOTARY)
5.  MOTHER NID/DEATH CERTIFICATE (TRANSLATION, NOTARY)
6.  FAMILY RELATIONSHIP CERTIFICATE (COLOUR PRINT+NOTARIZED)
7.  CERTIFICATE & TRANSCRIPT ATTESTED (E-APOSTILLE+NOTARIZED)
8.  BANK STATEMENT & SOLVENCY (SAVINGS ACCOUNT 23-25 LACS)
9.  FINANCIAL DOCUMENTS - TIN/TAX/TRADE LICENSE/EMPLOYMENT CERTIFICATE (COLOUR PRINT+NOTARIZED)
10. IELTS/TOPIK/TOEFL (COLOUR PRINT+NOTARIZED)
11. CV (CURRICULUM VITAE) (COLOUR PRINT+NOTARIZED)
12. RECOMMENDATION LETTER (COLOUR PRINT+NOTARIZED)
13. STATEMENT OF PURPOSE/SOP (COLOUR PRINT+NOTARIZED)

════════════════════════════════════════
DOCUMENT DETECTION KEYWORDS
════════════════════════════════════════
PASSPORT → passport no, date of issue, date of expiry, MRZ
STUDENT NID → national ID, NID No, voter ID, age 18-30
STUDENT BIRTH CERTIFICATE → birth registration, date of birth
FATHER/MOTHER NID → age 40+, "father of", "mother of"
FAMILY RELATIONSHIP CERTIFICATE / AFFIDAVIT → affidavit, relationship, notary public
CERTIFICATE & TRANSCRIPT → bachelor, master, CGPA, academic transcript
BANK STATEMENT & SOLVENCY → account statement, balance, 23 lac, 25 lac
FINANCIAL - TIN/TAX → TIN number, taxpayer, trade license
LANGUAGE PROFICIENCY → IELTS, TOEFL, TOPIK
CV → curriculum vitae, resume
RECOMMENDATION LETTER → recommend, pleased to recommend
STATEMENT OF PURPOSE / SOP → statement of purpose, my motivation, why Korea

════════════════════════════════════════
OUTPUT — always return an EXACT JSON ARRAY
════════════════════════════════════════
You MUST output a JSON ARRAY `[]` where each object is one identified document (or merged group). 
Do not miss any pages. 

[
  {{
    "start_page": 1,
    "end_page": 2,
    "program": "EAP/KLP or BACHELOR'S/MASTER'S",
    "official_pdf_name": "exact name from the correct set",
    "document_type": "plain english description",
    "confidence": 95,
    "trigger_keywords_found": ["word1", "word2"],
    "language_of_document": "English/Bengali/Mixed",
    "has_notary": true,
    "has_translation": true,
    "flag_for_manual_review": false,
    "flag_reason": null
  }}
]

════════════════════════════════════════
OCR TEXT TO ANALYZE:
════════════════════════════════════════
"""

        try:
            full_prompt = system_prompt + f"\n{ocr_text}\n"
            response = self.model.generate_content(full_prompt)
            
            text = response.text
            if "```json" in text:
                text = text.split("```json")[1].split("```")[0].strip()
            elif "```" in text:
                text = text.split("```")[1].split("```")[0].strip()
            
            classification_list = json.loads(text)
            
            # Ensure it's a list
            if isinstance(classification_list, dict):
                classification_list = [classification_list]
                
            return classification_list, "success"

        except Exception as e:
            return None, str(e)
