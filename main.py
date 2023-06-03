from flask import Flask, render_template, request
from translation import process_directory
import threading
import os
import time

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        source_dir = request.form.get('source_dir')
        dest_dir = request.form.get('dest_dir')
        start_file = request.form.get('start_file')
        convert_docs = request.form.get('convert_docs') == 'yes'
        override_docs = request.form.get('override_docs') == 'yes'
        convert_slides = request.form.get('convert_slides') == 'yes'
        override_slides = request.form.get('override_slides') == 'yes'
        copy_translate_others = request.form.get('copy_translate_others') == 'yes'
        override_others = request.form.get('override_others') == 'yes'

        if not all([source_dir, dest_dir]):
            return 'Both fields are required.'

        threading.Thread(target=run_translation, args=(source_dir, dest_dir, start_file, convert_docs, override_docs, convert_slides, override_slides, copy_translate_others, override_others)).start()

        return render_template('processing.html')

    return render_template('index.html')

def run_translation(source_dir, dest_dir, start_file, convert_docs, override_docs, convert_slides, override_slides, copy_translate_others, override_others):
    try:
        process_directory(source_dir, dest_dir, start_file, convert_docs, override_docs, convert_slides, override_slides, copy_translate_others, override_others)
        #time.sleep(1)  
    except Exception as e:
        print(f"An error occurred during translation: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True)
