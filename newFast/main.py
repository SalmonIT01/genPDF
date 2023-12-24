from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from io import BytesIO
from gen3 import*
import pdfkit



app = FastAPI()
@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.get("/items/{item_id}")
async def read_item(item_id: int):
    if item_id == 1:
        return {"message": "Sawadeekub"}
    if item_id == 2:
        return {"message": "Halo thailand"}
    return {"item_id": item_id}



config = pdfkit.configuration(wkhtmltopdf=r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")

@app.get("/genpdf")
async def generate_pdf():
    pdf_bytes = gen_word()
    # pdf_data = pdfkit.from_string(word.value, False, configuration=config)
    # os.remove('demo.pdf')
    return StreamingResponse(BytesIO(pdf_bytes), media_type="application/pdf")


