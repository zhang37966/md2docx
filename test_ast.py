import mistune
import sys

with open(sys.argv[1], 'r', encoding='utf-8') as f:
    text = f.read()

# Mock the re logic
import re
text = re.sub(r'([^\n])\n(#{1,6}\s+)', r'\1\n\n\2', text)
text = re.sub(r'(^|\n)(#{1,6}\s+.*?)\n([^\n])', r'\1\2\n\n\3', text)

ast = mistune.create_markdown(renderer='ast', plugins=['table', 'strikethrough'])(text)
for node in ast:
    print(node.get('type'))
