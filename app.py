from flask import Flask, render_template_string, send_from_directory
import gradio as gr
import threading
import socket
import easyocr
import pandas as pd
import tempfile
import os
import re

app = Flask(__name__)
gradio_port = None

reader = easyocr.Reader(['en'])

def ocr_app(image):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as temp:
        image.save(temp.name)
        image_path = temp.name

    results = reader.readtext(image_path, detail=0)
    lines = [line.strip() for line in results if line.strip()]

    extracted = {
        "College Name": "Not Found",
        "Student Name": "Not Found",
        "Reg No": "Not Found",
        "Semester": "Not Found",
        "Dept": "Not Found",
        "Year": "Not Found",
        "Staff Incharge": "Not Found"
    }

    for line in lines:
        l = line.lower()
        if "college" in l:
            extracted["College Name"] = re.sub(r".*college[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "name" in l:
            extracted["Student Name"] = re.sub(r".*name[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "reg" in l:
            extracted["Reg No"] = re.sub(r".*reg[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "semester" in l:
            extracted["Semester"] = re.sub(r".*semester[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "dept" in l or "department" in l or "cse" in l or "ece" in l or "eee" in l or "mech" in l:
            extracted["Dept"] = re.sub(r".*dept[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "year" in l:
            extracted["Year"] = re.sub(r".*year[^a-zA-Z0-9]*", "", line, flags=re.I).strip()
        elif "staff" in l or "incharge" in l:
            extracted["Staff Incharge"] = re.sub(r".*staff[^a-zA-Z0-9]*", "", line, flags=re.I).strip()

    df = pd.DataFrame([extracted])
    excel_path = os.path.join(tempfile.gettempdir(), "output.xlsx")
    df.to_excel(excel_path, index=False)

    full_text = "\n".join(lines)
    return full_text, excel_path

interface = gr.Interface(
    fn=ocr_app,
    inputs=gr.Image(type="pil"),
    outputs=["text", "file"],
    title="Handwritten OCR to Excel"
)

def get_free_port():
    s = socket.socket()
    s.bind(('', 0))
    port = s.getsockname()[1]
    s.close()
    return port

def run_gradio():
    global gradio_port
    gradio_port = get_free_port()
    interface.launch(share=False, server_port=gradio_port)

@app.route("/")
def index():
    if gradio_port:
        iframe_html = f'''
        <!DOCTYPE html>
        <html>
        <head>
          <title>OCR to Excel PWA</title>
          <link rel="manifest" href="/manifest.json">
          <meta name="theme-color" content="#3f51b5">
          <meta name="mobile-web-app-capable" content="yes">
          <meta name="apple-mobile-web-app-capable" content="yes">
        </head>
        <body style="margin:0">
          <iframe src="http://127.0.0.1:{gradio_port}" style="width:100%; height:100vh; border:none;"></iframe>
        </body>
        </html>
        '''
        return render_template_string(iframe_html)
    return "Gradio interface is not ready. Please refresh the page."

@app.route("/manifest.json")
def manifest():
    return send_from_directory(".", "manifest.json")

@app.route("/logo.png")
def logo():
    return send_from_directory("static", "logo.png")

if __name__ == "__main__":
    threading.Thread(target=run_gradio).start()
    app.run(debug=True)