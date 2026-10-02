"""
run_build.py — Builds the candidate DOCX using the modified builder in modifications/
"""
import time
import sys
from pathlib import Path

# Insert repository root and scripts directory
REPO_ROOT = Path(".").resolve()
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(REPO_ROOT / "scripts"))

# Load modified builder directly
import importlib.util
mod_path = Path("runs/DOC0000074469_20261002_1045/modifications/journal_article_builder.py").resolve()
spec = importlib.util.spec_from_file_location("modified_journal_builder", mod_path)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

from scripts.core.context import InputContext, BuildContext, InspectResult, ExtractResult

t0 = time.time()
pdf_path = Path("finished/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
if not pdf_path.exists():
    pdf_path = Path("input/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.pdf")
out_docx = Path("runs/DOC0000074469_20261002_1045/build/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx")
in_ctx = InputContext(pdf_path=pdf_path, output_docx=out_docx, config={}, temp_dir=Path("runs/DOC0000074469_20261002_1045"))
b_ctx = BuildContext(input_context=in_ctx, inspect_result=InspectResult(), extract_result=ExtractResult(), output_docx=out_docx)

builder = mod.JournalArticleBuilder()
res = builder.build(b_ctx)
elapsed = time.time() - t0
print(f"✅ Candidate DOCX built in {elapsed:.3f}s: {res}")
