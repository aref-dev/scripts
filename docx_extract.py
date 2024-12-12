from docx import Document
import re

# Load the document
def load_doc(filename):
    return Document(filename)

# Extract all text from the document
def get_text(doc):
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Extract specific section
def extract_section(text):
    # Pattern to match content between "TEXT A" and "TEXT B:"
    pattern = re.compile(r'TEXT A\s+(.*?)\s+TEXT B:', re.DOTALL)
    match = pattern.search(text)
    if match:
        section = match.group(1)  # Extract the section content
        # Split the section into lines and remove any empty lines
        lines = [line.strip() for line in section.splitlines() if line.strip()]
        return lines
    return []


def main(filename):
    doc = load_doc(filename)
    text = get_text(doc)
    data = extract_section(text)
    for line in data:
        print(line)


if __name__ == '__main__':
    main('FileName.docx')

