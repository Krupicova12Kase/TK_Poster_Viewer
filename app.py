from flask import Flask, flash, request, redirect, url_for, render_template, jsonify
import os
from os.path import join,dirname,realpath
from werkzeug.utils import secure_filename
app = Flask(__name__)

UPLOAD_FOLDER = join(dirname(realpath(__file__)), 'static/')
ALLOWED_EXTENSIONS = {"pptx"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

#display files
def display_files():
    files = []
    directory = join(dirname(realpath(__file__)), 'static/')
    try:
        x = 0
        for file in os.listdir(directory):
            filename = os.fsdecode(file)
            if filename.endswith(".pptx"): 
                x += 1
                files.append(filename)
        return ("SUCCESS", files)
    except:
        return ("FAIL", files)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

#upload
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            #file.filename = "test.pptx"
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            return redirect(url_for('upload_file', name=filename))
    display_files_output = display_files()
    if display_files_output[0] == "SUCCESS":
        return render_template("index.html", message="Test",files=display_files_output[1])
    else:
        return render_template("index.html", message="Test",files="Failed to load files")

@app.route('/spustit-python', methods=['POST'])
def spustit_python():
    import test
    test.slides_func()

    # Vrátíme JSON, aby JavaScript věděl, že se to povedlo
    return jsonify(status="success", message="Funkce proběhla na serveru.")
try:
    if __name__ == "__main__":
        app.run(debug=True)
except KeyboardInterrupt:
    print("App canceled")
except Exception as e:
    print(e)