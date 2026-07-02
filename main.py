from fastapi import FastAPI, UploadFile, File, HTTPException
from pyxlsb import open_workbook
import tempfile, os

app = FastAPI()

@app.get("/")
def home():
    return {"status": "ok"}

@app.post("/parse")
async def parse_xlsb(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsb"):
        raise HTTPException(status_code=400, detail="Only .xlsb files are allowed")

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsb")
    try:
        tmp.write(await file.read())
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
                        rows.append(dict(zip(header, values)))

        return {"rows": rows, "count": len(rows)}
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
