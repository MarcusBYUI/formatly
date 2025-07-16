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
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(MODEL_NAME)

def load_style_guide(style_name: str) -> dict:
    """Load the appropriate style guide based on the requested style name."""
    style_name = style_name.lower()
    return STYLE_GUIDES.get(style_name, STYLE_GUIDES["apa"])

def initialize_document_structure() -> dict:
    """Return a template for the document structure JSON."""
    return {
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

def create_detection_prompt(paragraphs: list) -> str:
    """Create the prompt for document structure detection."""
    json_structure = json.dumps(initialize_document_structure(), indent=4)
    return f"""
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
    print(prompt)

def parse_gemini_response(response_text: str) -> dict:
    """Parse the response from Gemini and return the structured data."""
    text = response_text.strip()
    
    # Remove markdown code fences if present
    if text.startswith("```json"):
        text = text[7:]
    elif text.startswith("```"):
        text = text[3:]
    
    if text.endswith("```"):
        text = text[:-3]
    
    # Remove any leading/trailing whitespace
    text = text.strip()
    
    # Validate that we have actual JSON content
    if not text or text[0] not in ['{', '[']:
        raise json.JSONDecodeError(f"Invalid JSON format: {text[:100]}...", text, 0)
    
    try:
        parsed_data = json.loads(text)
        return validate_document_structure(parsed_data)
    except json.JSONDecodeError as e:
        print(f"Warning: Failed to parse AI response as JSON: {e}")
        print(f"Raw response: {response_text[:200]}...")
        raise

def validate_document_structure(data: dict) -> dict:
    """Validate and sanitize the document structure returned by AI."""
    if not isinstance(data, dict):
        raise ValueError("Document structure must be a dictionary")
    
    # Ensure required keys exist with proper types
    validated = {
        "title_page": data.get("title_page", {}),
        "front_matter": data.get("front_matter", {}),
        "abstract": data.get("abstract", {}),
        "headings": data.get("headings", {"heading_1": [], "heading_2": [], "heading_3": []}),
        "body_content": data.get("body_content", {}),
        "special_elements": data.get("special_elements", {}),
        "citations_and_references": data.get("citations_and_references", {}),
        "tables_and_figures": data.get("tables_and_figures", {}),
        "appendices": data.get("appendices", {}),
        "back_matter": data.get("back_matter", {}),
        
        # Legacy keys for backward compatibility
        "block_quotes": data.get("block_quotes", []),
        "in_text_citations": data.get("in_text_citations", []),
        "tables": data.get("tables", []),
        "figures": data.get("figures", []),
        "footnotes": data.get("footnotes", []),
        "References": data.get("References", []),
        "references": data.get("references", [])  # Handle both cases
    }
    
    # Ensure lists are actually lists (legacy and new structure)
    list_keys = ["block_quotes", "in_text_citations", "tables", "figures", "footnotes", "References", "references"]
    for key in list_keys:
        if not isinstance(validated[key], list):
            validated[key] = []
    
    # Ensure nested dictionaries are dictionaries
    dict_keys = ["title_page", "front_matter", "abstract", "headings", "body_content", 
                 "special_elements", "citations_and_references", "tables_and_figures", 
                 "appendices", "back_matter"]
    for key in dict_keys:
        if not isinstance(validated[key], dict):
            validated[key] = {}
    
    return validated

def create_default_structure() -> dict:
    """Return a default document structure when parsing fails."""
    return {
        "title": "", "author": "", "affiliation": "", "abstract": [],
        "headings": {"heading_1": [], "heading_2": [], "heading_3": []},
        "block_quotes": [], "references": []
    }

def create_style_mappings() -> dict:
    """Return a dictionary of simple style mappings."""
    return {
        "title": "Title",
        "author": "Author",
        "affiliation": "Affiliation",
        "dedication": "Normal",
        "appendices_title": "Appendices Title",
        "appendix_title": "Appendix Title"
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
        self.style_name = style_name.lower()
        self.style_guide = load_style_guide(self.style_name)
        if not api_key:
            raise ValueError("API key for Gemini is required.")
        self.api_key = api_key
        self.model_name = model_name or "gemini-2.0-flash"  # Default fallback
        # Initialize rate limit manager
        self.rate_limit_manager = RateLimitManager(self.model_name)

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
                
                # Preserve the original outline level
                original_outline = style.paragraph_format.outline_level
                
                if 'font' in style_def:
                    p = doc.add_paragraph()
                    run = p.add_run()
                    self._apply_font_style(run.font, style_def['font'])
                    
                    font = style.font
                    for attr, value in style_def['font'].items():
                        if hasattr(font, attr):
                            setattr(font, attr, value)
                    
                    doc._body.remove(p._p)
                
                if 'paragraph' in style_def:
                    p = doc.add_paragraph()
                    self._apply_paragraph_style(p, style_def['paragraph'])
                    
                    p_format = style.paragraph_format
                    for attr, value in style_def['paragraph'].items():
                        if hasattr(p_format, attr):
                            setattr(p_format, attr, value)
                    
                    # Restore the outline level after applying other formatting
                    style.paragraph_format.outline_level = original_outline
                    doc._body.remove(p._p)
                
                style.hidden = False
                style.unhide_when_used = True
                style.quick_style = True
                
            except (KeyError, AttributeError):
                continue

    def format_document(self, input_path, output_path):
        doc = Document(input_path)

        self._customize_builtin_styles(doc)
        self._apply_margins(doc)
        self._add_page_numbers(doc)

        paragraphs_text = [p.text for p in doc.paragraphs]
        doc_structure = self._detect_paragraph_types(paragraphs_text)
        if not doc_structure.get("references"):
            self._find_references_manually(doc, doc_structure)

        self._format_content_in_place(doc, doc_structure)
        self._insert_page_breaks(doc)
        
        doc.save(output_path)

    def _detect_paragraph_types(self, paragraphs):
        try:
            model = configure_gemini(self.api_key)
            prompt = create_detection_prompt(paragraphs)
            
            def _make_request():
                """Internal function to make the actual API request."""
                response = model.generate_content(prompt)
                return parse_gemini_response(response.text)
            
            # Use rate limit manager to handle the request with automatic retries
            return self.rate_limit_manager.execute_with_rate_limit(_make_request)
        except json.JSONDecodeError:
            return create_default_structure()
        except Exception:
            return create_default_structure()

    def _format_content_in_place(self, doc, doc_structure):
        style_map = {}
        simple_style_map = create_style_mappings()
        section_map = create_section_mappings()

        # Apply simple style mappings
        for key, style in simple_style_map.items():
            text = doc_structure.get(key)
            if text and isinstance(text, str):
                style_map[text.strip()] = style

        # Apply section title mappings
        for key, title_text in section_map.items():
            if doc_structure.get(key):
                style_map[title_text] = "Heading 1"

        # Apply special section styles
        for key, style in [("abstract", "Normal"), ("block_quotes", "Block Quote"), ("references", "References")]:
            for text in doc_structure.get(key, []):
                if isinstance(text, str):
                    style_map[text.strip()] = style

        # Apply heading styles
        for level, heading_list in doc_structure.get("headings", {}).items():
            style_name = level.replace('_', ' ').title()
            for heading_obj in heading_list:
                if isinstance(heading_obj, dict):
                    heading_text = heading_obj.get('text', '').strip()
                    if heading_text:
                        style_map[heading_text] = style_name
                    for content_para in heading_obj.get('content', []):
                        if isinstance(content_para, str):
                            style_map[content_para.strip()] = "Normal"

        # Apply figure styles
        for fig in doc_structure.get("figures", []):
            if isinstance(fig, dict):
                number = fig.get("number", "").strip()
                caption = fig.get("caption", "").strip()
                if number:
                    style_map[number] = "Figure"
                if caption:
                    style_map[caption] = "Figure Caption"

        # Apply table styles
        for tbl in doc_structure.get("tables", []):
            if isinstance(tbl, dict):
                number = tbl.get("number", "").strip()
                title = tbl.get("title", "").strip()
                notes = tbl.get("notes", [])

                if number:
                    style_map[number] = "Table"
                if title:
                    style_map[title] = "Table Title"
                for note in notes:
                    if isinstance(note, str):
                        style_map[note.strip()] = "Table Note"

        # First pass: Apply heading styles to establish document structure
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            if text.lower() == "abstract" and text not in style_map:
                p.style = doc.styles["Heading 1"]
            elif text.lower() == "references" and text not in style_map:
                p.style = doc.styles["Heading 1"]
            elif text in style_map and style_map[text].startswith("Heading"):
                p.style = doc.styles[style_map[text]]

        # Second pass: Apply all other formatting
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            if text.lower() == "abstract" and text not in style_map:
                style_name = "Heading 1"
            elif text.lower() == "references" and text not in style_map:
                style_name = "Heading 1"
            else:
                style_name = style_map.get(text, "Normal")
            
            self._apply_advanced_style(p, style_name)

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
        config = self.style_guide["styles"].get(style_name, {})
        
        if "font" in config:
            for run in paragraph.runs:
                self._apply_font_style(run.font, config["font"])
        
        if "paragraph" in config:
            self._apply_paragraph_style(paragraph, config["paragraph"])

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