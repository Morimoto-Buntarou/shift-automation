from flask import Flask, render_template, request
import subprocess

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/execute', methods=['POST'])
def execute_script():
    if 'file' not in request.files:
        return render_template('index.html', message='No file part')

    file = request.files['file']

    if file.filename == '':
        return render_template('index.html', message='No selected file')

    file.save('uploaded_script.py')
    result = subprocess.run(['python', 'uploaded_script.py'], capture_output=True)

    return render_template('index.html', message=result.stdout.decode(), error=result.stderr.decode())

if __name__ == '__main__':
    app.run(debug=True)