from flask import Flask, request, send_file, render_template
import os
import uuid
import threading

def cleanup(path):
        import time
        time.sleep(10)
        if os.path.exists(path):
            os.remove(path)

print("Starting Flask app...")

app = Flask(__name__)

UPLOAD_FOLDER = "temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/download-template")
def download_template():
    from research_tool import create_input_template
    import uuid
    temp_path = f"temp/template_[{str(uuid.uuid4())[:8]}].xlsx"
    create_input_template(temp_path)
    
    response = send_file(temp_path, as_attachment=True, download_name="companies_input.xlsx")
    
    threading.Thread(target=cleanup, args=(temp_path,)).start()
    
    return response

@app.route("/analyze", methods=["POST"])
def analyze():
    if "file" not in request.files:
        return "No file uploaded", 400
    
    file = request.files["file"]

    if not file.filename.endswith(".xlsx"):
        return "Invalid file type. Please upload an .xlsx file.", 400
    session_id = str(uuid.uuid4())[:8]
    input_path = f"{UPLOAD_FOLDER}/input_{session_id}.xlsx"
    output_path = f"{UPLOAD_FOLDER}/output_{session_id}.xlsx"
    
    file.save(input_path)
    
    # Run analysis
    from research_tool import read_input, analyze_company, create_excel
    tickers = read_input(input_path)
    results = []
    for ticker in tickers:
        try:
            result = analyze_company(ticker)
            results.append(result)
        except Exception as e:
            print(f"Skipping {ticker} - {e}")
    
    create_excel(results, output_path)
    
    # Send file then delete both
    response = send_file(output_path, as_attachment=True, download_name=f"sec_analysis_{session_id}.xlsx")
    
    os.remove(input_path)

    threading.Thread(target=cleanup, args=(output_path,)).start()
    
    return response

if __name__ == "__main__":
    app.run(debug=True)