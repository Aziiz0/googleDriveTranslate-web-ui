from flask import Flask, render_template, request, redirect, url_for
from translation import process_directory
from traverse import traverse_directory
import threading
import os
import time
import configparser

try:
    config = configparser.ConfigParser()
    config.read('config.ini')
except Exception as e:
    print(f"Failed to read config file: {str(e)}")

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        source_dir = request.form.get('source_dir')
        if source_dir:
            return redirect(url_for('folders', root_dir=source_dir))
    return render_template('index.html')

@app.route('/folders', methods=['GET', 'POST'])
def folders():
    root_dir = request.args.get('root_dir')
    if request.method == 'POST':
        selected_folder_ids = request.form.getlist('folder_id')
        selected_folder_names = request.form.getlist('folder_name')
        selected_folders = list(zip(selected_folder_ids, selected_folder_names))
        dest_dir = request.form.get('dest_dir')
        if not dest_dir:
            return 'Destination directory is required.'
        threading.Thread(target=run_translation, args=(selected_folders, dest_dir)).start()
        return render_template('processing.html')
    else:
        folders = traverse_directory(root_dir)
        return render_template('folders.html', folders=folders, root_dir=root_dir)

def run_translation(folders, dest_dir):
    for folder in folders:
        folder_id, folder_name = folder
        try:
            process_directory(folder_id, dest_dir)
        except Exception as e:
            print(f"An error occurred during translation of {folder_name}: {str(e)}")

if __name__ == '__main__':
    app.run(debug=True)