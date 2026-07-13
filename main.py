from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from pyxlsb import open_workbook
from fastapi.middleware.cors import CORSMiddleware
import tempfile, os, base64

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://huron2--seenudev.sandbox.lightning.force.com"],  # or ["*"] while testing
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class ParseRequest(BaseModel):
    fileName: str
    fileBase64: str

@app.post("/parse")
async def parse_xlsb(payload: ParseRequest):
    if not payload.fileName.lower().endswith(".xlsb"):
        raise HTTPException(status_code=400, detail="Only .xlsb files are allowed")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb")
    try:
        file_bytes = base64.b64decode(payload.fileBase64)
        tmp.write(file_bytes)
        tmp.close()

        rows = []
        with open_workbook(tmp.name) as wb:
            with wb.get_sheet(1) as sheet:
                header = None
                for i, row in enumerate(sheet.rows()):
                    values = [c.v for c in row]
                    if i == 0:
                        header = values
                    else:
                        if header is None:
                            continue
                        rows.append(dict(zip(header, values)))

        return {"rows": rows, "count": len(rows)}

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Parse failed: {str(e)}")

    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
