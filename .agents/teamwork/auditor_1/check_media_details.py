import zipfile
from PIL import Image
import io

docx_path = r"output/DOC0000074469-Non-Parsable-en-US#FPREP_DXEGTK#.docx"

print("=== MEDIA FILES AND DIMENSIONS IN DOCX ===")
with zipfile.ZipFile(docx_path, "r") as zf:
    for info in zf.infolist():
        if info.filename.startswith("word/media/"):
            data = zf.read(info.filename)
            im = Image.open(io.BytesIO(data))
            print(f"{info.filename}: format={im.format}, size={im.size}, mode={im.mode}, file_size={info.file_size} bytes")
