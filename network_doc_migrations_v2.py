"""
Word Document Content Migration Script
Migrates content from old HLD/LLD Word documents to a new Word template
Preserves text, tables, and images/diagrams
"""

import os
import re
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from copy import deepcopy
import logging
from typing import Dict, List, Optional, Tuple
from io import BytesIO

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class WordDocumentMigrator:
    """Handles migration of content from old Word documents to new templates"""

    def __init__(self, old_doc_path: str, template_path: str, output_path: str):
        """
        Initialize the migrator with document paths

        Args:
            old_doc_path: Path to the old Word document
            template_path: Path to the new Word template
            output_path: Path for the output document
        """
        self.old_doc_path = old_doc_path
        self.template_path = template_path
        self.output_path = output_path
        self.old_doc = None
        self.new_doc = None
        self.section_mapping = {}

    def load_documents(self):
        """Load the old document and new template"""
        try:
            self.old_doc = Document(self.old_doc_path)
            self.new_doc = Document(self.template_path)
            logger.info(f"Successfully loaded old document: {self.old_doc_path}")
            logger.info(f"Successfully loaded template: {self.template_path}")
        except Exception as e:
            logger.error(f"Error loading documents: {str(e)}")
            raise

    def extract_sections(self, doc: Document) -> Dict[str, List]:
        """
        Extract sections and their content from a document

        Args:
            doc: The Word document to extract from

        Returns:
            Dictionary with section headings as keys and content lists as values
        """
        sections = {}
        current_section = None
        current_content = []

        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                para = None
                for p in doc.paragraphs:
                    if p._element == element:
                        para = p
                        break

                if para:
                    # Check if this is a heading
                    if para.style.name.startswith('Heading'):
                        # Save previous section if exists
                        if current_section:
                            sections[current_section] = current_content
                            current_content = []
                        current_section = para.text.strip()
                        logger.info(f"Found section: {current_section}")
                    elif current_section:
                        # Add paragraph to current section
                        current_content.append(('paragraph', para))

            elif element.tag.endswith('tbl'):  # Table
                if current_section:
                    for t in doc.tables:
                        if t._element == element:
                            current_content.append(('table', t))
                            logger.info(f"Found table in section: {current_section}")
                            break

        # Save last section
        if current_section:
            sections[current_section] = current_content

        return sections

    def find_matching_section(self, old_section: str, new_sections: List[str]) -> Optional[str]:
        """
        Find the best matching section in the new document

        Args:
            old_section: Section heading from old document
            new_sections: List of section headings in new document

        Returns:
            Best matching section heading or None
        """
        old_section_lower = old_section.lower()

        # First try exact match
        for new_section in new_sections:
            if new_section.lower() == old_section_lower:
                return new_section

        # Clean and split into words, removing common words that don't help matching
        stop_words = {'the', 'a', 'an', 'and', 'or', 'of', 'to', 'in', 'for', 'on', 'at', 'by'}

        # Get meaningful words from old section
        old_words = set(old_section_lower.split()) - stop_words
        old_words = {word.strip('.,;:!?()[]{}') for word in old_words if word.strip('.,;:!?()[]{})')}

        best_match = None
        best_match_count = 0

        # Try to find section with most words in common
        for new_section in new_sections:
            new_section_lower = new_section.lower()
            # Get meaningful words from new section
            new_words = set(new_section_lower.split()) - stop_words
            new_words = {word.strip('.,;:!?()[]{}') for word in new_words if word.strip('.,;:!?()[]{})')}

            # Find common words
            common_words = old_words & new_words

            # If at least 1 word matches, consider it a potential match
            if len(common_words) >= 1:
                if len(common_words) > best_match_count:
                    best_match = new_section
                    best_match_count = len(common_words)

        return best_match



    def insert_content_after_heading(self, heading_text: str, content_list: List[Tuple]):
        """
        Insert content after a specific heading in the new document

        Args:
            heading_text: The heading to insert content after
            content_list: List of (type, content) tuples
        """
        # Find the heading in the document element structure
        heading_element = None
        for element in self.new_doc.element.body:
            if element.tag.endswith('p'):
                for p in self.new_doc.paragraphs:
                    if p._element == element and p.text.strip() == heading_text:
                        heading_element = element
                        break
                if heading_element is not None:
                    break

        if heading_element is None:
            logger.warning(f"Could not find heading: {heading_text}")
            return

        # Get the parent (body) and the index of the heading
        body = self.new_doc.element.body
        heading_index = body.index(heading_element)
        insert_index = heading_index + 1

        # Track elements we're adding so we can insert them in order
        elements_to_insert = []

        for content_type, content in content_list:
            if content_type == 'paragraph':
                if content.text.strip() or self._has_images(content):  # Include if has text or images
                    # Create new paragraph with same content
                    new_para = self._create_paragraph_element(content)
                    elements_to_insert.append(new_para)

            elif content_type == 'table':
                # Create new table element
                new_table = self._copy_table_element(content)
                if new_table is not None:
                    elements_to_insert.append(new_table)

        # Insert all elements at the correct position
        for i, element in enumerate(elements_to_insert):
            body.insert(insert_index + i, element)

        logger.info(f"Inserted {len(elements_to_insert)} elements after section: {heading_text}")

    def _create_paragraph_element(self, source_para):
        """
        Create a new paragraph element with content from source paragraph

        Args:
            source_para: Source paragraph to copy

        Returns:
            New paragraph element
        """
        from docx.oxml import parse_xml
        from docx.oxml.ns import qn
        from docx.text.paragraph import Paragraph

        # Create new paragraph through the document
        new_p = self.new_doc.add_paragraph()

        # Clear default text
        new_p.clear()

        # Copy alignment if exists
        if source_para.alignment:
            new_p.alignment = source_para.alignment

        # Copy paragraph style if applicable
        if source_para.style and source_para.style.name != 'Normal':
            try:
                new_p.style = source_para.style.name
            except:
                pass

        # Copy runs with formatting
        for run in source_para.runs:
            new_run = new_p.add_run(run.text)

            # Copy formatting
            if run.bold:
                new_run.bold = True
            if run.italic:
                new_run.italic = True
            if run.underline:
                new_run.underline = True
            if run.font.name:
                new_run.font.name = run.font.name
            if run.font.size:
                new_run.font.size = run.font.size

        # Extract the element and remove from document body (we'll insert it at right place)
        element = new_p._element
        self.new_doc.element.body.remove(element)

        # Check for and copy images
        self._copy_images_from_paragraph(source_para, new_p)

        return element

    def _copy_table_element(self, source_table):
        """
        Create a copy of a table element

        Args:
            source_table: Source table to copy

        Returns:
            New table element
        """
        try:
            # Create new table with same dimensions
            new_table = self.new_doc.add_table(rows=len(source_table.rows),
                                              cols=len(source_table.columns))

            # Copy table style if available
            if source_table.style:
                try:
                    new_table.style = source_table.style
                except:
                    logger.debug("Could not copy table style")

            # Copy cell content
            for i, row in enumerate(source_table.rows):
                for j, cell in enumerate(row.cells):
                    new_cell = new_table.rows[i].cells[j]

                    # Clear default paragraph
                    if new_cell.paragraphs:
                        p = new_cell.paragraphs[0]
                        p.clear()

                    # Copy all paragraphs from source cell
                    for para_idx, para in enumerate(cell.paragraphs):
                        if para_idx == 0 and new_cell.paragraphs:
                            # Use existing first paragraph
                            new_para = new_cell.paragraphs[0]
                        else:
                            new_para = new_cell.add_paragraph()

                        # Copy text and formatting
                        for run in para.runs:
                            new_run = new_para.add_run(run.text)
                            if run.bold:
                                new_run.bold = True
                            if run.italic:
                                new_run.italic = True

            # Extract element and remove from body
            element = new_table._element
            self.new_doc.element.body.remove(element)

            return element

        except Exception as e:
            logger.error(f"Error copying table: {str(e)}")
            return None

    def _has_images(self, para):
        """
        Check if a paragraph contains images

        Args:
            para: Paragraph to check

        Returns:
            Boolean indicating if paragraph has images
        """
        for run in para.runs:
            if 'graphic' in run._element.xml or 'picture' in run._element.xml:
                return True
        return False

    def _copy_images_from_paragraph(self, source_para, target_para):
        """
        Copy images from source paragraph to target paragraph

        Args:
            source_para: Source paragraph with potential images
            target_para: Target paragraph to add images to
        """
        try:
            # Check each run for inline shapes/images
            for run in source_para.runs:
                if 'graphic' in run._element.xml:
                    # Try to extract and copy the image
                    for inline in run._element.iter():
                        if 'blip' in inline.tag:
                            for attr in inline.attrib:
                                if 'embed' in attr:
                                    rId = inline.attrib[attr]
                                    try:
                                        image_part = source_para.part.related_parts[rId]
                                        image_data = image_part.blob

                                        # Add image to target paragraph
                                        new_run = target_para.add_run()
                                        new_run.add_picture(BytesIO(image_data), width=Inches(4))
                                        logger.debug("Successfully copied an image")
                                    except Exception as e:
                                        logger.warning(f"Could not copy image: {str(e)}")
        except Exception as e:
            logger.debug(f"Error checking for images: {str(e)}")

    def setup_section_mapping(self, custom_mapping: Dict[str, str] = None):
        """
        Setup mapping between old and new document sections

        Args:
            custom_mapping: Optional custom mapping dictionary
        """
        if custom_mapping:
            self.section_mapping = custom_mapping
            print("\n=== CUSTOM SECTION MAPPING APPLIED ===")
            for old_section, new_section in custom_mapping.items():
                print(f"  '{old_section}' -> '{new_section}'")
        else:
            # Auto-detect mappings
            old_sections = self.extract_sections(self.old_doc)
            new_sections = self.extract_sections(self.new_doc)

            print("\n=== AUTO-DETECTING SECTION MAPPINGS ===")
            print("Looking for sections with at least 1 word in common...\n")

            # Define stop words to ignore in matching
            stop_words = {'the', 'a', 'an', 'and', 'or', 'of', 'to', 'in', 'for', 'on', 'at', 'by'}

            for old_section in old_sections.keys():
                matching = self.find_matching_section(old_section, list(new_sections.keys()))
                if matching:
                    self.section_mapping[old_section] = matching

                    # Show which words matched
                    old_words = set(old_section.lower().split()) - stop_words
                    old_words = {word.strip('.,;:!?()[]{}') for word in old_words if word.strip('.,;:!?()[]{})')}
                    new_words = set(matching.lower().split()) - stop_words
                    new_words = {word.strip('.,;:!?()[]{}') for word in new_words if word.strip('.,;:!?()[]{})')}
                    common_words = old_words & new_words

                    print(f"✓ MAPPED: '{old_section}'")
                    print(f"      TO: '{matching}'")
                    print(f"  COMMON: {', '.join(common_words)}")
                    print()

                    logger.info(f"Mapped: '{old_section}' -> '{matching}' (common words: {common_words})")
                else:
                    print(f"✗ NO MATCH: '{old_section}'")
                    print()
                    logger.warning(f"No matching section found for: '{old_section}'")

            # Summary
            print(f"\n=== MAPPING SUMMARY ===")
            print(f"Total sections in old document: {len(old_sections)}")
            print(f"Successfully mapped: {len(self.section_mapping)}")
            print(f"Unmapped sections: {len(old_sections) - len(self.section_mapping)}")

            if self.section_mapping:
                print("\n=== FINAL MAPPING TABLE ===")
                for old, new in self.section_mapping.items():
                    print(f"  '{old}' -> '{new}'")

    def migrate_content(self):
        """Main method to perform the content migration"""
        logger.info("Starting content migration...")

        # Load documents
        self.load_documents()

        # Extract sections from old document
        old_sections = self.extract_sections(self.old_doc)
        logger.info(f"Found {len(old_sections)} sections in old document")

        # Setup section mappings if not provided
        if not self.section_mapping:
            self.setup_section_mapping()

        # Migrate content based on mapping
        for old_section, new_section in self.section_mapping.items():
            if old_section in old_sections:
                content = old_sections[old_section]
                self.insert_content_after_heading(new_section, content)
            else:
                logger.warning(f"Section '{old_section}' not found in old document")

        # Save the new document
        self.new_doc.save(self.output_path)
        logger.info(f"Migration complete. Output saved to: {self.output_path}")

    def print_section_summary(self):
        """Print a summary of sections found in both documents"""
        if not self.old_doc or not self.new_doc:
            self.load_documents()

        old_sections = self.extract_sections(self.old_doc)
        new_sections = self.extract_sections(self.new_doc)

        print("\n=== OLD DOCUMENT SECTIONS ===")
        for i, section in enumerate(old_sections.keys(), 1):
            content_count = len(old_sections[section])
            print(f"{i}. {section} ({content_count} items)")

        print("\n=== NEW TEMPLATE SECTIONS ===")
        for i, section in enumerate(new_sections.keys(), 1):
            print(f"{i}. {section}")

        print("\n=== SECTION MAPPING ===")
        if self.section_mapping:
            for old, new in self.section_mapping.items():
                print(f"'{old}' -> '{new}'")
        else:
            print("No mapping defined yet")


def main():
    """Main function to run the Word Document Migrator with command line arguments"""

    import argparse
    import sys
    import json

    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Migrate content from old Word HLD/LLD documents to new templates",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Basic usage:
    python migrate_word.py old_doc.docx template.docx output.docx
  
  With analysis only (no migration):
    python migrate_word.py old_doc.docx template.docx output.docx --analyze
  
  With custom mapping file:
    python migrate_word.py old_doc.docx template.docx output.docx --mapping mapping.json
  
  Interactive mode (review mappings before migration):
    python migrate_word.py old_doc.docx template.docx output.docx --interactive
        """
    )

    # Required arguments
    parser.add_argument('old_doc',
                       help='Path to the old Word document (source)')
    parser.add_argument('template',
                       help='Path to the new Word template')
    parser.add_argument('output',
                       help='Path for the output document')

    # Optional arguments
    parser.add_argument('--analyze', '-a',
                       action='store_true',
                       help='Only analyze documents and show sections without migrating')
    parser.add_argument('--mapping', '-m',
                       help='Path to JSON file containing custom section mappings')
    parser.add_argument('--interactive', '-i',
                       action='store_true',
                       help='Review auto-detected mappings before migration')
    parser.add_argument('--verbose', '-v',
                       action='store_true',
                       help='Enable verbose logging')

    # Parse arguments
    args = parser.parse_args()

    # Configure logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    # Load custom mapping if provided
    custom_mapping = {}
    if args.mapping:
        try:
            with open(args.mapping, 'r') as f:
                custom_mapping = json.load(f)
                print(f"Loaded custom mapping from: {args.mapping}")
        except FileNotFoundError:
            print(f"Error: Mapping file not found: {args.mapping}")
            sys.exit(1)
        except json.JSONDecodeError:
            print(f"Error: Invalid JSON in mapping file: {args.mapping}")
            sys.exit(1)

    try:
        # Create migrator instance
        migrator = WordDocumentMigrator(args.old_doc, args.template, args.output)

        # Load documents
        migrator.load_documents()

        # Show section analysis
        print("\nAnalyzing documents...")
        migrator.print_section_summary()

        # If analyze only mode, exit here
        if args.analyze:
            print("\n[Analysis complete - no migration performed]")
            sys.exit(0)

        # Setup mapping
        if custom_mapping:
            migrator.setup_section_mapping(custom_mapping)
        else:
            # Auto-detect mappings
            migrator.setup_section_mapping()

        # Interactive mode - allow user to review mappings
        if args.interactive and not custom_mapping:
            print("\n" + "="*50)
            response = input("Do you want to proceed with these mappings? (yes/no/edit): ").lower().strip()

            if response == 'edit':
                print("\nEnter custom mappings (or press Enter to skip):")
                print("Format: old_section -> new_section")
                print("Type 'done' when finished\n")

                custom_mapping = {}
                while True:
                    mapping_input = input("Mapping: ").strip()
                    if mapping_input.lower() == 'done':
                        break
                    if ' -> ' in mapping_input:
                        old, new = mapping_input.split(' -> ', 1)
                        custom_mapping[old.strip()] = new.strip()
                        print(f"  Added: {old.strip()} -> {new.strip()}")
                    elif mapping_input:
                        print("  Invalid format. Use: old_section -> new_section")

                if custom_mapping:
                    migrator.setup_section_mapping(custom_mapping)

            elif response != 'yes':
                print("Migration cancelled.")
                sys.exit(0)

        # Perform migration
        print("\nStarting migration...")
        migrator.migrate_content()

        print(f"\n✅ Success! Migrated content saved to: {args.output}")

        # Option to save the mapping for future use
        if not custom_mapping and migrator.section_mapping:
            save_mapping = input("\nSave these mappings for future use? (yes/no): ").lower().strip()
            if save_mapping == 'yes':
                mapping_file = args.output.replace('.docx', '_mapping.json')
                with open(mapping_file, 'w') as f:
                    json.dump(migrator.section_mapping, f, indent=2)
                print(f"Mappings saved to: {mapping_file}")

    except FileNotFoundError as e:
        print(f"\n❌ Error: Could not find file - {e}")
        sys.exit(1)
    except PermissionError as e:
        print(f"\n❌ Error: Permission denied - {e}")
        print("Make sure the files are not open in Word and you have write permissions.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error during migration: {e}")
        if args.verbose:
            logger.error(f"Migration failed: {str(e)}", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()