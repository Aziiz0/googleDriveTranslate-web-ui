# googleDriveTranslate-web-ui

This project requires Python 3.11.3 or newer.

## Installation

1. Clone the repository by running the following command in your terminal:
    ```
    git clone https://github.com/Aziiz0/googleDriveTranslate-web-ui.git
    ```
    
2. Navigate into the project directory:
    ```
    cd googleDriveTranslate-web-ui
    ```

3. Open `environment.env` file and replace `C:\Python\python.exe` with the path to your own Python executable file:
    ```
    PYTHON=<PATH_TO_YOUR_PYTHON_EXECUTABLE>
    ```
    Replace `<PATH_TO_YOUR_PYTHON_EXECUTABLE>` with the path to your Python executable. For example, `C:\Users\Username\AppData\Local\Programs\Python\Python311\python.exe`.

4. Run the installation script:
    ```
    install.bat
    ```

5. Setup your Google API and get the `.json` file containing your API key. You can follow the instructions in this [link](https://developers.google.com/workspace/guides/create-project).

6. Once you have your API key, place it in the main directory of `googleDriveTranslate-web-ui` and rename it to `googleKey.json`.

7. Finally, you can start the application by running:
    ```
    web-ui.bat
    ```
