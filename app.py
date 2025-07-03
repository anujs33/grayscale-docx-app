import os
import requests
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
CLOUDCONVERT_API_KEY = os.environ.get("CLOUDCONVERT_API_KEY")

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("docxfile")
        if not file or not file.filename.endswith(".docx"):
            return "Please upload a valid DOCX file."

        filename = secure_filename(file.filename)
        local_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(local_path)

        pdf_url = convert_to_grayscale_pdf(local_path)
        if not pdf_url:
            return "Conversion failed. Please try again."

        pdf_response = requests.get(pdf_url)
        pdf_path = os.path.join(RESULT_FOLDER, filename.replace(".docx", ".pdf"))
        with open(pdf_path, "wb") as f:
            f.write(pdf_response.content)

        return send_file(pdf_path, as_attachment=True)

    return render_template("index.html")

def convert_to_grayscale_pdf(local_file_path):
    import_task = requests.post(
        "https://api.cloudconvert.com/v2/import/upload",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}
    ).json()["data"]

    upload_url = import_task["result"]["form"]["url"]
    upload_params = import_task["result"]["form"]["parameters"]

    with open(local_file_path, "rb") as f:
        files = {"file": f}
        upload = requests.post(upload_url, data=upload_params, files=files)

    if upload.status_code != 204:
        return None

    import_id = import_task["id"]

    convert_response = requests.post(
        "https://api.cloudconvert.com/v2/tasks",
        headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}", "Content-Type": "application/json"},
        json={
            "task": {
                "operation": "convert",
                "input": import_id,
                "input_format": "docx",
                "output_format": "pdf",
                "engine": "office",
                "output": {
                    "pdf_color": "gray"
                }
            }
        }
    ).json()

    convert_task_id = convert_response["data"]["id"]
    status_url = f"https://api.cloudconvert.com/v2/tasks/{convert_task_id}"

    while True:
        status = requests.get(status_url, headers={"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}).json()
        if status["data"]["status"] in ["finished", "error"]:
            break

    if status["data"]["status"] != "finished":
        return None

    export_task = status["data"]["result"]["files"][0]
    return export_task["url"]

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)