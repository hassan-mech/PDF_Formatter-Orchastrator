import re
from pathlib import Path

builder_path = Path("scripts/stages/builders/journal_article_builder.py")
code = builder_path.read_text(encoding="utf-8")

# Record original in modifications/
mod_dir = Path("runs/DOC0000074469_20261002_1045/modifications")
mod_dir.mkdir(parents=True, exist_ok=True)
(mod_dir / "journal_article_builder_v1.py").write_text(code, encoding="utf-8")

# Let's inspect modifications needed
print("Original code length:", len(code))
