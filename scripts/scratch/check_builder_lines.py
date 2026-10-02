import sys
from pathlib import Path

# Read existing builder to get all text constants and imports
builder_path = Path("scripts/stages/builders/journal_article_builder.py")

# Let's inspect the exact lines where _create_layout_row, _add_section_heading, etc. are defined
lines = builder_path.read_text(encoding="utf-8").splitlines()
print(f"Total lines in current builder: {len(lines)}")
