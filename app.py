from flask import Flask, flash, request, redirect, url_for, render_template, jsonify
import os
from os.path import join,dirname,realpath
from werkzeug.utils import secure_filename
from time import sleep

UPLOAD_FOLDER = join(dirname(realpath(__file__)), 'static/')
ALLOWED_EXTENSIONS = {"pptx"}
REQUEST_NAMES = ["hlavicka","leva","stred","prava"]

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
    print("function")
    if request.method == 'POST':
        print("method")
        # check if the post request has the file part
        for name in REQUEST_NAMES:
            print("loop " + str(name))
            if name not in request.files:
                flash(f"Request part {name} is not in the provided request!")
                print(f"Request part {name} is not in the provided request!")
                return redirect(request.url)
            file = request.files[name]
            # If the user does not select a file, the browser submits an empty file without a filename.
            if file.filename == '':
                flash('No selected file')
                print('No selected file')
                return redirect(request.url)
            file.filename = name + ".pptx"

            print("idk1")

            if file and allowed_file(file.filename):

                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                print("Saved")   
        #Slide conversion
        sleep(2)
        import converter
        converter.slides_func()
        return redirect(url_for('upload_file', name=filename))  

    display_files_output = display_files()
    print("displaying files")       
    if display_files_output[0] == "SUCCESS":
        print("returning")
        return render_template("index.html", message="Test",files=display_files_output[1])
    else:
        return render_template("index.html", message="Test",files="Failed to load files")
    

if __name__ == "__main__":
    app.run(debug=True)