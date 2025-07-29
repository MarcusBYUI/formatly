from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml.shared import OxmlElement
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE, WD_BUILTIN_STYLE
from style_guides import STYLE_GUIDES
from spell_check import DocumentChecker, format_error_report
from utils.auto_corrector import AutoCorrector
from utils.gemini_corrector import GeminiCorrector
from utils.formatting_analyzer import FormattingAnalyzer
from utils.batch_processor import BatchProcessor
from utils.rate_limit_manager import RateLimitManager
from utils.dynamic_chunk_calculator import DynamicChunkCalculator
import google.generativeai as genai
import json
import os
import re
from dotenv import load_dotenv
from pathlib import Path

# Load environment variables
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-2.0-flash-exp")
def configure_gemini(api_key: str) -> genai.GenerativeModel:
    """Configure and return the Gemini model."""
    print(f"[DEBUG] configure_gemini: Starting with API key length: {len(api_key) if api_key else 0}")
    print(f"[DEBUG] configure_gemini: Model name: {MODEL_NAME}")
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(MODEL_NAME)
    print(f"[DEBUG] configure_gemini: Successfully created model instance")
    return model

def load_style_guide(style_name: str) -> dict:
    """Load the appropriate style guide based on the requested style name."""
    print(f"[DEBUG] load_style_guide: Requested style: {style_name}")
    style_name = style_name.lower()
    print(f"[DEBUG] load_style_guide: Normalized style name: {style_name}")
    style_guide = STYLE_GUIDES.get(style_name, STYLE_GUIDES["apa"])
    print(f"[DEBUG] load_style_guide: Found style guide: {'Yes' if style_name in STYLE_GUIDES else 'No, using APA fallback'}")
    return style_guide

def initialize_document_structure() -> dict:
    """Return a template for the document structure JSON."""
    print(f"[DEBUG] initialize_document_structure: Creating document structure template")
    structure = {
        "title_page": {
            "title": "",
            "author": "",
            "institution": "",
            "course": "",
            "instructor": "",
            "due_date": ""
        },
        "front_matter": {
            "dedication": "",
            "acknowledgments": "",
            "preface": "",
            "table_of_contents": "",
            "list_of_figures": "",
            "list_of_tables": "",
            "list_of_abbreviations": "",
            "glossary": ""
        },
        "abstract": {
            "text": "",
            "keywords": []
        },
        "headings": {
            "heading_1": [{"text": "", "content": []}],
            "heading_2": [{"text": "", "content": []}],
            "heading_3": [{"text": "", "content": []}],
            "heading_4": [{"text": "", "content": []}],
            "heading_5": [{"text": "", "content": []}]
        },
        "body_content": {
            "introduction": [],
            "methodology": [],
            "results": [],
            "discussion": [],
            "conclusion": [],
            "body_paragraphs": []
        },
        "special_elements": {
            "block_quotes": [{"quote": "", "source": "", "page": ""}],
            "equations": [{"equation": "", "number": ""}],
            "code_blocks": [{"code": "", "language": ""}],
            "lists": [{"type": "ordered|unordered", "items": []}]
        },
        "citations_and_references": {
            "in_text_citations": [{"citation": "", "type": "parenthetical|narrative"}],
            "footnotes": [{"text": "", "number": ""}],
            "endnotes": [{"text": "", "number": ""}],
            "references": [{"entry": "", "type": "book|journal|website|other"}]
        },
        "tables_and_figures": {
            "tables": [{"number": "", "title": "", "content": "", "notes": []}],
            "figures": [{"number": "", "caption": "", "content": "", "source": ""}],
            "charts": [{"number": "", "title": "", "content": ""}],
            "images": [{"number": "", "caption": "", "alt_text": ""}]
        },
        "appendices": {
            "appendix_sections": [{"label": "", "title": "", "content": []}],
            "supplementary_materials": [{"title": "", "content": ""}]
        },
        "back_matter": {
            "bibliography": [],
            "index": [],
            "author_bio": ""
        }
    }
    print(f"[DEBUG] initialize_document_structure: Template created with {len(structure)} main sections")
    return structure

def create_detection_prompt(paragraphs: list) -> str:
    """Create the prompt for document structure detection."""
    print(f"[DEBUG] create_detection_prompt: Creating prompt for {len(paragraphs)} paragraphs")
    json_structure = json.dumps(initialize_document_structure(), indent=4)
    print(f"[DEBUG] create_detection_prompt: JSON structure template size: {len(json_structure)} chars")
    prompt = f"""
    You are an AI processor for academic documents. Your task is to classify and extract the content below into a structured JSON format that reflects all components requiring formatting.

    Return only valid JSON containing these keys without placeholders or sample text:
    
    {json_structure}
    
    ⛔ CONSTRAINTS:
    - Maintain original text order
    - Recognize all sections including dedications, acknowledgments, table of contents
    - Empty sections must return empty containers ([] or {{}})
    - Include appendices, tables, footnotes only if present in input
    - Ensure valid JSON output without explanations or markdown

    Input content (paragraphs):

    {'\n'.join(paragraphs)}

    Return valid JSON only without explanation or markdown.
    """
    print(f"[DEBUG] create_detection_prompt: Prompt created, length: {len(prompt)} chars")
    return prompt

def parse_gemini_response(response_text: str) -> dict:
    """
    Parse and validate the response from Gemini in a single pass.
    
    Args:
        response_text: Raw response text from Gemini API
        
    Returns:
        dict: Validated and sanitized document structure
        
    Raises:
        ValueError: If the response cannot be parsed as valid JSON
    """
    if not response_text or not isinstance(response_text, str):
        raise ValueError("Empty or invalid response text")
        
    # Clean and prepare the response text
    text = response_text.strip()
    
    # Remove markdown code fences if present
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    text = text.rstrip("`").strip()
    
    if not text:
        raise ValueError("Empty response after cleaning")
    
    try:
        # Parse JSON in a single pass
        data = json.loads(text)
        print(f"{data}")
        return validate_document_structure(data)
    except json.JSONDecodeError as e:
        error_msg = f"Failed to parse JSON response: {str(e)}"
        print(f"[ERROR] {error_msg}")
        # Try to find the start of the JSON object
        json_start = text.find('{')
        if json_start != -1:
            text = text[json_start:]
            try:
                data = json.loads(text)
                return validate_document_structure(data)
            except json.JSONDecodeError:
                pass  # Re-raise the original error if still failing
        raise ValueError(error_msg) from e

def validate_document_structure(data: dict) -> dict:
    """
    Validate and sanitize the document structure with minimal processing.
    
    Args:
        data: Raw parsed JSON data from Gemini
        
    Returns:
        dict: Validated and sanitized document structure
        
    Raises:
        ValueError: If the data structure is invalid
    """
    if not isinstance(data, dict):
        raise ValueError("Document structure must be a dictionary")
    
    # Define the schema for the document structure
    schema = {
        "title_page": dict,
        "front_matter": dict,
        "abstract": dict,
        "headings": dict,
        "body_content": dict,
        "special_elements": dict,
        "citations_and_references": dict,
        "tables_and_figures": dict,
        "appendices": dict,
        "back_matter": dict,
        # Legacy fields
        "block_quotes": list,
        "in_text_citations": list,
        "tables": list,
        "figures": list,
        "footnotes": list,
        "References": list,
        "references": list
    }
    
    # Initialize with defaults
    validated = {}
    
    # Validate and set each field
    for key, expected_type in schema.items():
        value = data.get(key)
        
        # Set default value if not present or wrong type
        if value is None or not isinstance(value, expected_type):
            validated[key] = [] if expected_type == list else {}
        else:
            validated[key] = value
    
    # Ensure headings has required sub-keys
    if not all(k in validated["headings"] for k in ["heading_1", "heading_2", "heading_3"]):
        validated["headings"] = {
            "heading_1": [],
            "heading_2": [],
            "heading_3": []
        }
    
    return validated

def create_default_structure() -> dict:
    """Return a default document structure when parsing fails."""
    return {
        "Title": "", "author": "", "affiliation": "", "abstract": [],
        "headings": {"heading_1": [], "heading_2": [], "heading_3": []},
        "block_quotes": [], "references": []
    }

def create_style_mappings() -> dict:
    """Return a comprehensive dictionary of style mappings for all document elements."""
    return {
        # Title page
        "title": "Title",
        "author": "Author",
        "affiliation": "Affiliation",
        "institution": "Institution",
        "course": "Course",
        "instructor": "Instructor",
        "due_date": "Date",
        
        # Front matter
        "dedication": "Dedication",
        "acknowledgments": "Acknowledgments",
        "preface": "Preface",
        "table_of_contents": "TOC Heading",
        "list_of_figures": "List of Figures",
        "list_of_tables": "List of Tables",
        "list_of_abbreviations": "List of Abbreviations",
        
        # Main content
        "abstract": "Abstract",
        "keywords": "Keywords",
        "heading_1": "Heading 1",
        "heading_2": "Heading 2",
        "heading_3": "Heading 3",
        "heading_4": "Heading 4",
        "heading_5": "Heading 5",
        "paragraph": "Normal",
        "block_quote": "Block Quote",
        "figure": "Figure",
        "table": "Table",
        "footnote": "Footnote",
        "endnote": "Endnote",

        # Body content (ensure strict mapping for all AI structure keys)
        "introduction": "Normal",
        "methodology": "Normal",
        "results": "Normal",
        "discussion": "Normal",
        "conclusion": "Normal",
        "body_paragraphs": "Normal",

        # Appendices and Figures
        "appendix": "Appendix",
        "appendix_heading": "Appendix Heading",
        "figure": "Figure",
        "figure_caption": "Figure Caption",
        
        # References
        "references": "References",
        "bibliography": "Bibliography",
        
        # Appendices
        "appendix": "Appendix",
        "appendix_heading": "Appendix Heading",
        
        # Back matter
        "glossary": "Glossary",
        "index": "Index",
        "author_bio": "Author Bio"
    }

def create_section_mappings() -> dict:
    """Return a dictionary of section title mappings."""
    return {
        "acknowledgements": "Acknowledgements",
        "table_of_contents": "Table of Contents",
        "list_of_figures": "List of Figures",
        "list_of_tables": "List of Tables",
        "preface": "Preface",
        "glossary": "Glossary",
    }

class AdvancedFormatter:
    def __init__(self, style_name, api_key: str = API_KEY, model_name: str = MODEL_NAME):
        print(f"[DEBUG] Initializing AdvancedFormatter with style: {style_name}")
        self.style_name = style_name.lower()
        self.style_guide = load_style_guide(self.style_name)
        if not api_key:
            raise ValueError("API key for Gemini is required.")
        self.api_key = api_key
        self.model_name = model_name or "gemini-2.0-flash"  # Default fallback
        print(f"[DEBUG] Using model: {self.model_name}")
        # Initialize rate limit manager
        self.rate_limit_manager = RateLimitManager(self.model_name)
        print(f"[DEBUG] AdvancedFormatter initialized successfully")

    def _customize_builtin_styles(self, doc):
        """Customize built-in styles while preserving their outline levels."""
        heading_map = {
            'Heading 1': 'Heading 1',
            'Heading 2': 'Heading 2',
            'Heading 3': 'Heading 3',
            'Heading 4': 'Heading 4',
            'Heading 5': 'Heading 5',
            'Heading 6': 'Heading 6',
            'Heading 7': 'Heading 7',
            'Heading 8': 'Heading 8',
            'Heading 9': 'Heading 9'
        }
        
        for builtin_name, style_name in heading_map.items():
            try:
                style = doc.styles[builtin_name]
                style_def = self.style_guide['styles'].get(style_name, {})
                print(f"[DEBUG] Customizing style: {builtin_name} -> {style_name}")
                # Preserve the original outline level
                original_outline = style.paragraph_format.outline_level
                
                # Apply font styles to the built-in style
                if 'font' in style_def:
                    p = doc.add_paragraph()
                    run = p.add_run()
                    self._apply_font_style(run.font, style_def['font'])
                    
                    # Update the built-in style's font
                    font = style.font
                    for attr, value in style_def['font'].items():
                        if hasattr(font, attr):
                            setattr(font, attr, value)
                    
                    # Remove the temporary paragraph
                    doc._body.remove(p._p)
                
                if 'paragraph' in style_def:
                    # Apply paragraph styles to the built-in style
                    # by creating a temporary paragraph, applying the style
                    # to it, and then extracting the desired attributes
                    # from the temporary paragraph's format
                    p = doc.add_paragraph()
                    self._apply_paragraph_style(p, style_def['paragraph'])
                    
                    # Copy over the desired attributes from the temporary
                    # paragraph's format to the built-in style's format
                    p_format = style.paragraph_format
                    for attr, value in style_def['paragraph'].items():
                        if hasattr(p_format, attr):
                            setattr(p_format, attr, getattr(p.paragraph_format, attr))
                    
                    # Restore the outline level after applying other formatting
                    style.paragraph_format.outline_level = original_outline
                    doc._body.remove(p._p)
                
                # Ensure the style is visible in the Styles pane
                style.hidden = False
                
                # Auto-show the style in the Styles pane when it is used
                style.unhide_when_used = True
                
                # Allow the style to be applied with the Quick Style button
                style.quick_style = True
                
            except (KeyError, AttributeError):
                continue

    def format_document(self, input_path, output_path):
        print(f"[DEBUG] Starting document formatting: {input_path} -> {output_path}")
        doc = Document(input_path)
        print(f"[DEBUG] Document loaded, paragraph count: {len(doc.paragraphs)}")

        print(f"[DEBUG] Customizing built-in styles...")
        self._customize_builtin_styles(doc)
        print(f"[DEBUG] Applying margins...")
        self._apply_margins(doc)
        print(f"[DEBUG] Adding page numbers...")
        self._add_page_numbers(doc)

        paragraphs_text = [p.text for p in doc.paragraphs]
        print(f"[DEBUG] Extracting paragraphs: {len(paragraphs_text)} total")
        print(f"[DEBUG] Detecting paragraph types using AI...")
        doc_structure = self._detect_paragraph_types(paragraphs_text)
        print(f"[DEBUG] Document structure detected")
        
        if not doc_structure.get("references"):
            print(f"[DEBUG] No references found in structure, searching manually...")
            self._find_references_manually(doc, doc_structure)
        else:
            print(f"[DEBUG] References found in structure: {len(doc_structure.get('references', []))}")

        print(f"[DEBUG] Formatting content in place...")
        self._format_content_in_place(doc, doc_structure)
        print(f"[DEBUG] Inserting page breaks...")
        self._insert_page_breaks(doc)
        
        print(f"[DEBUG] Saving formatted document...")
        doc.save(output_path)
        print(f"[DEBUG] Document formatting completed successfully")

    def _detect_paragraph_types(self, paragraphs):
        print(f"[DEBUG] Detecting paragraph types for {len(paragraphs)} paragraphs")
        try:
            print(f"[DEBUG] Configuring Gemini model for paragraph detection...")
            model = configure_gemini(self.api_key)
            print(f"[DEBUG] Creating detection prompt...")
            prompt = create_detection_prompt(paragraphs)
            print(f"[DEBUG] Prompt created, length: {len(prompt)} chars")
            
            def _make_request():
                """Internal function to make the actual API request."""
                print(f"[DEBUG] Making API request to Gemini...")
                response = model.generate_content(prompt)
                print(f"[DEBUG] Received response from Gemini")
                print(f"[DEBUG] Raw Gemini API response: {response.text}")
                print(f"[DEBUG] Response length: {len(response.text)} chars")
                print(f"[DEBUG] Starting response parsing...")
                return parse_gemini_response(response.text)
            
            # Use rate limit manager to handle the request with automatic retries
            print(f"[DEBUG] Executing request with rate limit manager...")
            result = self.rate_limit_manager.execute_with_rate_limit(_make_request)
            print(f"[DEBUG] Successfully detected paragraph types")
            return result
        except json.JSONDecodeError as e:
            print(f"[DEBUG] JSON decode error in paragraph detection: {e}")
            print(f"[DEBUG] Falling back to default structure")
            return create_default_structure()
        except Exception as e:
            print(f"[DEBUG] Exception in paragraph detection: {e}")
            print(f"[DEBUG] Falling back to default structure")
            return create_default_structure()

    def _format_content_in_place(self, doc, doc_structure):
        # Create a style map to look up which paragraph text corresponds to which style
        style_map = {}
        simple_style_map = create_style_mappings()
        section_map = create_section_mappings()

        # Ensure all required styles exist in the DOCX
        from docx.enum.style import WD_STYLE_TYPE
        needed_styles = set(simple_style_map.values()) | {"References", "Keywords", "Block Quote", "Table Title", "Figure Caption", "Appendix Heading", "Normal"}
        for style_name in needed_styles:
            if style_name not in doc.styles:
                doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)

        # Handle title page elements
        title_page = doc_structure.get("title_page", {})
        for key, value in title_page.items():
            if key in simple_style_map and value and isinstance(value, str):
                # Map paragraph text to the corresponding style
                style_map[value.strip()] = simple_style_map[key]

        # Handle front matter elements
        front_matter = doc_structure.get("front_matter", {})
        for key, value in front_matter.items():
            if key in simple_style_map and value and isinstance(value, str):
                # Map paragraph text to the corresponding style
                style_map[value.strip()] = simple_style_map[key]

        # Handle abstract
        abstract = doc_structure.get("abstract", {})
        if abstract.get("text"):
            # Map the abstract text to the Abstract style
            style_map[abstract["text"].strip()] = "Abstract"
        if abstract.get("keywords"):
            for keyword in abstract["keywords"]:
                if isinstance(keyword, str):
                    # Map each keyword to the Keywords style
                    style_map[keyword.strip()] = "Keywords"

        # Handle headings
        headings = doc_structure.get("headings", {})
        for level, heading_list in headings.items():
            if level in simple_style_map:
                for heading_obj in heading_list:
                    if isinstance(heading_obj, dict) and heading_obj.get("text"):
                        # Map the heading text to the corresponding style
                        style_map[heading_obj["text"].strip()] = simple_style_map[level]
                        # Map content paragraphs under this heading to 'Normal' style
                        if heading_obj.get("content") and isinstance(heading_obj["content"], list):
                            for para in heading_obj["content"]:
                                if isinstance(para, str):
                                    style_map[para.strip()] = "Normal"
                    elif isinstance(heading_obj, str):
                        # Map the heading text to the corresponding style
                        style_map[heading_obj.strip()] = simple_style_map[level]

        # Handle body content
        body_content = doc_structure.get("body_content", {})
        for section, content_list in body_content.items():
            if section in simple_style_map and isinstance(content_list, list):
                for content in content_list:
                    if isinstance(content, str):
                        # Map the body content text to the corresponding style
                        style_map[content.strip()] = simple_style_map[section]

        # Handle special elements
        special_elements = doc_structure.get("special_elements", {})
        block_quotes = special_elements.get("block_quotes", [])
        for quote_obj in block_quotes:
            if isinstance(quote_obj, dict) and quote_obj.get("quote"):
                # Map the block quote text to the Block Quote style
                style_map[quote_obj["quote"].strip()] = "Block Quote"

        # Handle citations and references
        citations_refs = doc_structure.get("citations_and_references", {})
        references = citations_refs.get("references", [])
        for ref_obj in references:
            if isinstance(ref_obj, dict) and ref_obj.get("entry"):
                # Map the reference text to the References style
                style_map[ref_obj["entry"].strip()] = "References"
            elif isinstance(ref_obj, str):
                # Map the reference text to the References style
                style_map[ref_obj.strip()] = "References"

        footnotes = citations_refs.get("footnotes", [])
        for footnote_obj in footnotes:
            if isinstance(footnote_obj, dict) and footnote_obj.get("text"):
                # Map the footnote text to the Footnote style
                style_map[footnote_obj["text"].strip()] = "Footnote"

        # Handle tables and figures
        tables_figures = doc_structure.get("tables_and_figures", {})
        
        tables = tables_figures.get("tables", [])
        for table_obj in tables:
            if isinstance(table_obj, dict):
                if table_obj.get("title"):
                    # Map the table title to the Table Title style
                    style_map[table_obj["title"].strip()] = "Table Title"
                if table_obj.get("number"):
                    # Map the table number to the Table style
                    style_map[table_obj["number"].strip()] = "Table"

        figures = tables_figures.get("figures", [])
        for figure_obj in figures:
            if isinstance(figure_obj, dict):
                if figure_obj.get("caption"):
                    # Map the figure caption to the Figure Caption style
                    style_map[figure_obj["caption"].strip()] = "Figure Caption"
                if figure_obj.get("number"):
                    # Map the figure number to the Figure style
                    style_map[figure_obj["number"].strip()] = "Figure"
                # Map any figure label or title
                if figure_obj.get("label"):
                    style_map[figure_obj["label"].strip()] = "Figure"
                if figure_obj.get("title"):
                    style_map[figure_obj["title"].strip()] = "Figure Caption"

        # Handle appendices
        appendices = doc_structure.get("appendices", {})
        appendix_sections = appendices.get("appendix_sections", [])
        for appendix_obj in appendix_sections:
            if isinstance(appendix_obj, dict):
                if appendix_obj.get("title"):
                    # Map the appendix title to the Appendix Heading style
                    style_map[appendix_obj["title"].strip()] = "Appendix Heading"
                if appendix_obj.get("label"):
                    # Map the appendix label to the Appendix style
                    style_map[appendix_obj["label"].strip()] = "Appendix"
                # Map any appendix content as Normal
                if appendix_obj.get("content") and isinstance(appendix_obj["content"], list):
                    for para in appendix_obj["content"]:
                        if isinstance(para, str):
                            style_map[para.strip()] = "Normal"

        # Handle back matter
        back_matter = doc_structure.get("back_matter", {})
        for key, value in back_matter.items():
            if key in simple_style_map:
                if isinstance(value, str) and value:
                    # Map the back matter text to the corresponding style
                    style_map[value.strip()] = simple_style_map[key]
                elif isinstance(value, list):
                    for item in value:
                        if isinstance(item, str):
                            # Map the back matter text to the corresponding style
                            style_map[item.strip()] = simple_style_map[key]

        # Apply section title mappings
        # Iterate through the section_map dictionary
        for key, title_text in section_map.items():
            # Check if the key exists in the doc_structure dictionary
            if doc_structure.get(key):
                # Map the section title to the Heading 1 style
                # in the style_map dictionary
                style_map[title_text] = "Heading 1"

        # --- Preprocessing: split paragraphs with multiple mapped elements ---
        # Build a set of all mapped element texts (for fast lookup)
        mapped_texts = set(style_map.keys())
        # Sort mapped_texts by length descending to match longest first
        mapped_texts_sorted = sorted(mapped_texts, key=len, reverse=True)

        # Gather all reference entries for robust splitting
        reference_entries = set()
        citations_refs = doc_structure.get("citations_and_references", {})
        references = citations_refs.get("references", [])
        for ref_obj in references:
            if isinstance(ref_obj, dict) and ref_obj.get("entry"):
                reference_entries.add(ref_obj["entry"].strip())
            elif isinstance(ref_obj, str):
                reference_entries.add(ref_obj.strip())

        # Work on a copy since we'll be modifying the paragraphs list
        i = 0
        while i < len(doc.paragraphs):
            p = doc.paragraphs[i]
            text = p.text.strip()
            if not text:
                i += 1
                continue
            # --- Special: Split 'References' + reference entry in one paragraph ---
            if text.startswith("References"):
                for ref in reference_entries:
                    idx = text.find(ref)
                    if idx > 0:
                        before = text[:idx].strip()
                        after = text[idx:].strip()
                        p.text = before
                        new_p = doc.add_paragraph(after)
                        p._element.addnext(new_p._element)
                        # Restart processing for this index (in case after also needs splitting)
                        break
            matches = [mt for mt in mapped_texts_sorted if mt and text.startswith(mt)]
            if len(matches) > 1:
                # Paragraph starts with more than one mapped element, ambiguous, skip for now
                i += 1
                continue
            elif len(matches) == 1 and text != matches[0]:
                # Paragraph starts with a mapped element but contains more text (merged)
                matched = matches[0]
                rest = text[len(matched):].lstrip('\n').lstrip()
                # Replace with two paragraphs: mapped element and rest
                p.text = matched
                if rest:
                    new_p = doc.add_paragraph(rest)
                    # Insert new_p after p
                    p._element.addnext(new_p._element)
                # Restart processing for this index (in case rest also needs splitting)
                continue
            i += 1

        # --- Diagnostic pass: write mapping status for each paragraph to file ---
        with open('paragraph_mapping_diagnostics.txt', 'w', encoding='utf-8') as diag_file:
            diag_file.write("[DEBUG] Paragraph mapping diagnostics:\n")
            for p in doc.paragraphs:
                text = p.text.strip()
                if not text:
                    continue
                preview = text[:100] + ("..." if len(text) > 100 else "")
                if text in style_map:
                    diag_file.write(f"[MATCH] Exact: '{preview}' → '{style_map[text]}'\n")
                else:
                    # Check for substring match (both directions) for body styles
                    body_styles = {v for k, v in simple_style_map.items() if k in ["introduction", "methodology", "results", "discussion", "conclusion", "body_paragraphs"]}
                    found = False
                    for key, style in style_map.items():
                        if style in body_styles and (text in key or key in text):
                            diag_file.write(f"[MATCH] Substring: '{preview}' ≈ '{key[:100]}' → '{style}'\n")
                            found = True
                            break
                    if not found:
                        # Try robust reference matching
                        reference_style = simple_style_map.get("references", "References")
                        def normalize(s):
                            return ' '.join(s.lower().split())
                        norm_text = normalize(text)
                        for key, style in style_map.items():
                            if style == reference_style and normalize(key) == norm_text:
                                diag_file.write(f"[MATCH] Reference (normalized): '{preview}' → '{style}'\n")
                                found = True
                                break
                        if not found:
                            diag_file.write(f"[UNMATCHED] '{preview}'\n")

        # First pass: Apply all styles from the style map with exact matching
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            # Get style name from map with exact matching
            style_name = style_map.get(text)
            
            if style_name is not None:
                # Apply the style if it exists in the document
                if style_name in doc.styles:
                    p.style = style_name
                    continue
                else:
                    print(f"[DEBUG] Style '{style_name}' is not defined in the document. Available styles: {list(doc.styles)}")
                    raise ValueError(f"Style '{style_name}' is not defined in the document")
            
            # Special handling for section titles that must be exact matches
            if text in ["Abstract", "References"]:
                if text in doc.styles:
                    p.style = text
                else:
                    # Fall back to Heading 1 for standard section titles if exact style not found
                    p.style = "Heading 1"

            # If still no style, try substring/fuzzy match for body content
            else:
                preview = text[:100] + ("..." if len(text) > 100 else "")
                from difflib import get_close_matches
                close_matches = get_close_matches(text, list(style_map.keys()), n=3, cutoff=0.3)
                # Only allow substring match for known body content styles
                body_styles = {v for k, v in simple_style_map.items() if k in ["introduction", "methodology", "results", "discussion", "conclusion", "body_paragraphs"]}
                found = False
                for key, style in style_map.items():
                    if style in body_styles and (text in key or key in text):
                        if style in doc.styles:
                            p.style = style
                            found = True
                            print(f"[DEBUG] Applied body style '{style}' to paragraph (substring match): '{preview}'")
                            break
                if found:
                    continue
                # Robust reference mapping (normalize whitespace/case)
                reference_style = simple_style_map.get("references", "References")
                def normalize(s):
                    return ' '.join(s.lower().split())
                norm_text = normalize(text)
                for key, style in style_map.items():
                    if style == reference_style and normalize(key) == norm_text:
                        if style in doc.styles:
                            p.style = style
                            found = True
                            print(f"[DEBUG] Applied reference style '{style}' to paragraph (normalized match): '{preview}'")
                            break
                if found:
                    continue
                # Fallback: assign 'Normal' style to unmatched body/reference paragraphs, warn user
                heading_styles = {v for k, v in simple_style_map.items() if k.startswith("heading_") or k in ["title", "abstract", "references"]}
                # Fallback for figure captions/labels
                if text.lower().startswith("figure") and "Figure" in doc.styles:
                    p.style = "Figure"
                    print(f"[WARNING] Unmatched paragraph starting with 'Figure' assigned 'Figure' style: '{preview}'")
                    continue
                # Fallback for appendix headings/labels
                if "appendix" in text.lower() and "Appendix" in doc.styles:
                    p.style = "Appendix"
                    print(f"[WARNING] Unmatched paragraph containing 'appendix' assigned 'Appendix' style: '{preview}'")
                    continue
                if not any(text.lower().startswith(h.lower()) for h in heading_styles):
                    if "Normal" in doc.styles:
                        p.style = "Normal"
                        print(f"[WARNING] Unmatched paragraph assigned 'Normal' style: '{preview}'")
                        continue
                print(f"[DEBUG] No style mapping found for paragraph (preview): '{preview}'")
                print(f"[DEBUG] Closest style map keys: {close_matches}")
                print(f"[DEBUG] All style map keys: {list(style_map.keys())}")
                raise ValueError(f"No style mapping found for element: '{preview}'. All document elements must have a defined style.")

        # Second pass: Apply advanced formatting
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            # Get style name with exact matching
            style_name = style_map.get(text)
            
            if style_name is None:
                # Only allow specific section titles without explicit mapping
                if text in ["Abstract", "References"]:
                    style_name = text
                else:
                    raise ValueError(f"No style mapping found for element: '{text}'. "
                                   "All document elements must have a defined style.")
            
            # Apply the style with strict validation
            self._apply_advanced_style(p, style_name)
            
            # Special handling for references
            if style_name == "References":
                self._apply_reference_formatting(p, style=self.style_name)

    def _apply_margins(self, doc):
        section = doc.sections[0]
        margin_map = {
            "left": "left_margin", "right": "right_margin",
            "top": "top_margin", "bottom": "bottom_margin",
            "header": "header_distance", "footer": "footer_distance",
            "gutter": "gutter"
        }
        for key, value in self.style_guide["margins"].items():
            if key in margin_map:
                setattr(section, margin_map[key], value)

    def _find_references_manually(self, doc, doc_structure):
        reference_started = False
        references = []
        for para in doc.paragraphs:
            text = para.text.strip()
            if reference_started:
                if text:
                    references.append(text)
            elif text.lower() == "references":
                reference_started = True
        if references:
            doc_structure["references"] = references

    def _add_page_numbers(self, doc):
        position = self.style_guide["meta"].get("page_numbers")
        if not position:
            return

        for section in doc.sections:
            if position == "header-right":
                header = section.header
                paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                return

            run = paragraph.add_run()
            fldChar_begin = OxmlElement('w:fldChar')
            fldChar_begin.set(qn('w:fldCharType'), 'begin')
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = "PAGE"
            fldChar_end = OxmlElement('w:fldChar')
            fldChar_end.set(qn('w:fldCharType'), 'end')
            run._r.extend([fldChar_begin, instrText, fldChar_end])

    def _apply_reference_formatting(self, paragraph, style="apa"):
        text = paragraph.text

        for run in list(paragraph.runs):
            run._element.getparent().remove(run._element)

        def add_run(run_text, italic=False):
            run = paragraph.add_run(run_text)
            run.italic = italic
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if "*" in text:
            parts = re.split(r"\*(.*?)\*", text)
            for idx, part in enumerate(parts):
                add_run(part, italic=bool(idx % 2))
            return

        style = style.lower()
        italic_segment = ""
        if style in ("apa", "harvard"):
            m = re.search(r"\)\.\s*([^\.]+)", text)
            if m:
                italic_segment = m.group(1).strip()
        elif style == "mla":
            m = re.search(r"\.\s*([^\.]+)\.", text)
            if m:
                italic_segment = m.group(1).strip()

        if italic_segment and italic_segment in text:
            pre, _, post = text.partition(italic_segment)
            add_run(pre, italic=False)
            add_run(italic_segment, italic=True)
            add_run(post, italic=False)
        else:
            add_run(text, italic=False)

    def _apply_font_style(self, font, font_config):
        for attr, value in font_config.items():
            if attr == 'color':
                if isinstance(value, str):
                    font.color.rgb = RGBColor.from_string(value)
                else:
                    font.color.rgb = value
            elif hasattr(font, attr):
                setattr(font, attr, value)

    def _apply_paragraph_style(self, paragraph, para_config):
        p_format = paragraph.paragraph_format
        for attr, value in para_config.items():
            if hasattr(p_format, attr):
                setattr(p_format, attr, value)

    def _apply_advanced_style(self, paragraph, style_name):
        """Apply advanced styling to a paragraph with strict style validation."""
        try:
            # Try to get the style from the document
            style = paragraph.style = style_name
        except KeyError:
            # If style doesn't exist, raise an error
            raise ValueError(
                f"Style '{style_name}' is not defined in the document. "
                "Please ensure all required styles are defined in the style guide."
            )
        
        # Apply style settings from style guide if available
        style_config = self.style_guide["styles"].get(style_name, {})
        
        if "font" in style_config:
            for run in paragraph.runs:
                self._apply_font_style(run.font, style_config["font"])
        
        if "paragraph" in style_config:
            self._apply_paragraph_style(paragraph, style_config["paragraph"])

    def _insert_page_breaks(self, doc):
        page_break_styles = ["Heading 1", "Appendices Title"]

        for para in doc.paragraphs:
            if para.style.name.split('.')[0] in page_break_styles:
                if para._p.getprevious() is not None:
                    br = OxmlElement('w:br')
                    br.set(qn('w:type'), 'page')

def parse_arguments():
    """Parse and return command line arguments."""
    import argparse
    parser = argparse.ArgumentParser(description="Format academic documents with consistent styling.", formatter_class=argparse.RawTextHelpFormatter)
    
    # Re-add all existing arguments exactly as before
    parser.add_argument("input", nargs="?", 
                       help="Input Word document (required if not using --interactive). If only filename is given, looks in 'documents' folder.")
    parser.add_argument("-o", "--output", help="Output file path (default: saves to 'formatting' folder with <input>_formatted_<style>.docx)")
    parser.add_argument("-s", "--style", default="apa", choices=["apa", "mla", "chicago"],
                       help="Citation style (default: apa)")
    parser.add_argument("-e", "--english", default="american", choices=["american", "british", "australian", "canadian"],
                       help="English variation (default: american)")
    parser.add_argument("-i", "--interactive", action="store_true",
                       help="Interactive mode - prompts for input file if not provided")
    parser.add_argument("-l", "--list-styles", action="store_true",
                       help="List available citation styles and exit")
    # New mutually-exclusive main actions
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--report-only', action='store_true', help='Generate reports only')
    group.add_argument('--fix-spelling', action='store_true', help='Fix spelling, punctuation, and capitalization')
    group.add_argument('--format', action='store_true', help='Format document')
    group.add_argument('--fix-errors', action='store_true', help='Fix errors and format document')
    group.add_argument('--batch-process', action='store_true', help='Process document using Gemini Batch API')
    
    # Add chunk size parameter for API request chunking
    parser.add_argument('--chunk-size', type=int, default=3, 
                       help='Number of paragraphs to process in each API request (default: 3)')
    
    # Add batch mode parameters
    parser.add_argument('--batch', action='store_true',
                       help='Use Batch Mode for processing')
    parser.add_argument('--max-wait', type=int, default=3600,
                       help='Maximum time to wait for batch processing in seconds')
    parser.add_argument('--no-wait', action='store_true',
                       help='Submit batch job but don\'t wait for completion')
    
    return parser.parse_args()

def validate_style(style_name: str) -> str:
    """Validate and return the style name, defaulting to APA if invalid."""
    return style_name if style_name.lower() in STYLE_GUIDES else "apa"

def determine_output_path(input_path: Path, output_arg: str, style_name: str) -> Path:
    """Determine the output path based on input and arguments."""
    if output_arg:
        return Path(output_arg)
    
    # Default to saving in 'formatting' folder
    formatting_dir = Path("formatting")
    formatting_dir.mkdir(exist_ok=True)
    return formatting_dir / f"{input_path.stem}_formatted_{style_name}{input_path.suffix}"

def list_available_styles():
    """List all available citation styles."""
    print("Available citation styles:")
    for style_name, style_config in STYLE_GUIDES.items():
        meta = style_config.get("meta", {})
        print(f"  {style_name.upper():8} - {meta.get('default_font', 'Times New Roman')} font")
        print(f"          {'Title page: ' + ('Yes' if meta.get('title_page') else 'No'):20}")
        print(f"          {'Abstract: ' + ('Required' if meta.get('requires_abstract') else 'Optional'):20}")
        print()

def get_input_file_interactive():
    """Prompt user for input file in interactive mode."""
    print("\nNote: If you provide just a filename, it will be searched in the 'documents' folder.")
    while True:
        input_file = input("Enter the path to your Word document: ").strip()
        if not input_file:
            print("Please provide a file path.")
            continue
        
        input_path = Path(input_file)
        
        # Check if file exists as provided
        if input_path.exists():
            if input_path.suffix.lower() == '.docx':
                return input_path
            else:
                print("Error: File must be a Word document (.docx)")
        else:
            # Check in documents folder
            documents_path = Path("documents") / input_path.name
            if documents_path.exists():
                if documents_path.suffix.lower() == '.docx':
                    return documents_path
                else:
                    print("Error: File must be a Word document (.docx)")
            else:
                print(f"Error: File '{input_file}' not found (also checked in 'documents' folder).")

def validate_input_file(input_path: str) -> Path:
    """Validate input file exists and is a Word document."""
    if not input_path:
        raise ValueError("Input file path is required")
    
    path = Path(input_path)
    
    # If the file doesn't exist as provided, check in the documents folder
    if not path.exists():
        documents_path = Path("documents") / path.name
        if documents_path.exists():
            path = documents_path
        else:
            raise FileNotFoundError(f"Input file not found: {input_path} (also checked in 'documents' folder)")
    
    if path.suffix.lower() != '.docx':
        raise ValueError(f"Input file must be a Word document (.docx), got: {path.suffix}")
    
    return path

def main():
    args = parse_arguments()
    
    # Handle list styles option
    if args.list_styles:
        list_available_styles()
        return 0
    
    # Handle input file
    input_path = None
    if args.input:
        try:
            input_path = validate_input_file(args.input)
        except (FileNotFoundError, ValueError) as e:
            print(f"Error: {e}")
            return 1
    elif args.interactive:
        try:
            input_path = get_input_file_interactive()
        except KeyboardInterrupt:
            print("\nOperation cancelled by user.")
            return 1
    else:
        print("Error: Input file is required. Use --interactive mode or provide a file path.")
        print("Usage: python app.py input.docx [options]")
        print("   or: python app.py --interactive [options]")
        return 1
    
    style_name = validate_style(args.style)
    
    try:
        # Load document for analysis
        doc = Document(str(input_path))
        paragraphs_text = [p.text for p in doc.paragraphs if p.text.strip()]
        
        # Handle new mutually exclusive processing modes
        if args.report_only:
            # Generate comprehensive report that analyzes but doesn't modify document
            print(f"📊 Generating comprehensive report for: {input_path}")
            checker = DocumentChecker(english_variation=args.english)
            formatter_analyzer = FormattingAnalyzer(style_name)
            try:
                # Generate spelling report
                report = checker.get_correction_report(paragraphs_text)
                print("\n=== SPELLING REPORT ===")
                print(format_error_report(report))
                
                # Generate formatting analysis report
                print("\n=== FORMATTING ANALYSIS ===")
                analysis_result = formatter_analyzer.analyze_document(str(input_path))
                analysis_report = formatter_analyzer.generate_report(analysis_result)
                print(analysis_report)
                
                print("\n✅ Report generated successfully. No changes made to document.")
                return 0
            except Exception as e:
                print(f"Error during report generation: {e}")
                return 1
            finally:
                checker.close()
        
        elif args.fix_errors:
            # Automatically fix spelling errors before formatting
            print(f"🔧 Auto-fixing spelling errors: {input_path}")
            auto_corrector = AutoCorrector(english_variation=args.english)
            gemini_corrector = None
            try:
                print("\n🔍 Initializing Gemini AI for combined spelling, punctuation, and capitalization correction...")
                gemini_corrector = GeminiCorrector()
                print("Gemini AI initialized successfully for comprehensive text correction.")
            except Exception as e:
                print(f"Warning: Could not initialize Gemini corrector: {e}")
                print("Falling back to basic spell checker.")

            try:
                # Use GeminiCorrector if available
                if gemini_corrector:
                    print("\n🔧 Using Gemini AI to apply comprehensive text corrections...")
                    print("   Simultaneous correction of spelling, punctuation, and capitalization in progress...")
                    corrected_paragraphs, correction_summary = gemini_corrector.correct_paragraphs(paragraphs_text, chunk_size=args.chunk_size)
                else:
                    # Fall back to existing AutoCorrector
                    print("\n🔧 Applying basic automatic corrections...")
                    print("   Note: Only basic spelling will be checked without Gemini AI")
                    corrected_paragraphs, correction_summary = auto_corrector.apply_all(paragraphs_text, chunk_size=args.chunk_size)
                
                # Update paragraphs in the document with corrected content
                non_empty_paragraphs = [p for p in doc.paragraphs if p.text.strip()]
                idx = 0
                for i, p in enumerate(doc.paragraphs):
                    if p.text.strip() and idx < len(corrected_paragraphs):
                        p.text = corrected_paragraphs[idx]
                        idx += 1
                
                # Display correction summary
                print(auto_corrector.generate_correction_explanation(correction_summary))
                print(f"✅ Applied {correction_summary['total_corrections']} corrections")
                
                if correction_summary['total_corrections'] > 0:
                    # Get breakdown by type
                    by_type = correction_summary.get("corrections_by_type", {})
                    print(f"✅ Successfully applied corrections using Gemini AI:")
                    print(f"   🔤 Spelling: {by_type.get('spelling', 0)}")
                    print(f"   🔣 Punctuation: {by_type.get('punctuation', 0)}")
                    print(f"   🔠 Capitalization: {by_type.get('capitalization', 0)}")
                else:
                    print("✅ No issues found that needed correction.")
                    
                print("\n📝 Proceeding with document formatting...")
            except Exception as e:
                print(f"Error during auto-fix: {e}")
                return 1
            finally:
                if auto_corrector:
                    auto_corrector.close()
                if gemini_corrector:
                    del gemini_corrector
        
        elif args.format:
            # TODO: Implement formatting-only mode that skips all spelling/grammar checks
            # This flag should skip all spellcheck operations and proceed directly to formatting
            print(f"🎨 Applying formatting only (skipping spelling): {input_path}")
            
            # TODO: Add logic to bypass all spelling checking phases
            # - Skip DocumentChecker initialization
            # - Skip error reporting
            # - Skip user prompts about errors
            # - Proceed directly to formatting operations
            
            pass  # Placeholder - formatting will happen in the main flow below
            
        elif args.batch_process:
            print(f"🚀 Processing document in batch mode: {input_path}")
            batch_processor = BatchProcessor()
            try:
                # Create a batch job for the document
                job_result = batch_processor.process_document_batch(
                    document_text="\n\n".join(paragraphs_text),
                    style_name=style_name,
                    wait_for_completion=not args.no_wait,
                    max_wait_time=args.max_wait
                )

                if job_result['completed']:
                    output_path = determine_output_path(input_path, args.output, style_name)
                    output_path.parent.mkdir(parents=True, exist_ok=True)

                    with open(output_path, 'w', encoding='utf-8') as f:
                        f.write(job_result['document'])

                    print(f"✓ Batch processing completed. Document saved as {output_path}")
                else:
                    print(f"Batch job {job_result['job_name']} submitted. Check status later.")

                return 0
            except Exception as e:
                print(f"Error during batch processing: {e}")
                return 1
            finally:
                batch_processor.cleanup()
        
        else:
            # Default mode - spell check then format
            print(f"🔍 Checking spelling: {input_path}")
            checker = DocumentChecker(english_variation=args.english)
            try:
                report = checker.get_correction_report(paragraphs_text)
                print(format_error_report(report))
                
                # Ask user if they want to continue with formatting
                if report['total_spelling_errors'] > 0:
                    continue_format = input("\nFound spelling errors. Continue with formatting? (y/N): ")
                    if continue_format.lower() not in ['y', 'yes']:
                        print("Formatting cancelled. Please fix errors and try again.")
                        return 1
                print()
            finally:
                checker.close()
        
        # Only proceed with formatting if we're not in report-only mode
        # TODO: Enhance conditional logic to handle all flag combinations properly
        if not args.report_only:
            # Check API key for formatting
            if not API_KEY:
                print("Error: GEMINI_API_KEY environment variable not set.")
                print("Please add your Gemini API key to the .env file.")
                return 1
            
            print(f"📝 Formatting document: {input_path}")
            print(f"🎨 Style: {style_name.upper()}, English: {args.english.capitalize()}")
            
            formatter = AdvancedFormatter(style_name)
            output_path = determine_output_path(input_path, args.output, style_name)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Check if output file already exists
            if output_path.exists():
                response = input(f"Output file '{output_path}' already exists. Overwrite? (y/N): ")
                if response.lower() not in ['y', 'yes']:
                    print("Operation cancelled.")
                    return 1
            
            # Handle flag-specific formatting behavior
            if args.fix_errors:
                # For --fix-errors: We've already applied corrections to the document
                # Save the document with corrections first
                doc_with_corrections_path = input_path.parent / f"{input_path.stem}_corrected{input_path.suffix}"
                doc.save(str(doc_with_corrections_path))
                
                # Then format the corrected document
                formatter.format_document(str(doc_with_corrections_path), str(output_path))
                
                # Clean up the temporary file
                import os
                try:
                    os.remove(str(doc_with_corrections_path))
                except Exception as e:
                    print(f"Note: Could not remove temporary file {doc_with_corrections_path}: {e}")
            else:
                # For other modes: Apply normal formatting workflow
                formatter.format_document(str(input_path), str(output_path))
            print(f"✓ Document formatted in {style_name.upper()} style and saved as {output_path}")
        
        return 0
        
    except FileNotFoundError as e:
        print(f"Error: Input file not found: {e.filename}")
        return 1
    except PermissionError as e:
        print(f"Error: Permission denied when accessing file: {e.filename}")
        print("Make sure the file is not open in another application.")
        return 1
    except ValueError as e:
        print(f"Error: {e}")
        return 1
    except Exception as e:
        print(f"Error: Failed to format document: {e}")
        print("Please check your input file and try again.")
        return 1

if __name__ == "__main__":
    raise SystemExit(main())