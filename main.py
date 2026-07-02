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
        content = await file.read()
        tmp.write(content)
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
