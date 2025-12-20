# Voice Control for windows

Welcome! This is a simple application that lets you control your Windows computer using your voice. You can open programs, change the volume, and even shut down or restart your computer, all by speaking to it.

### How to Use the Application

The application listens for a special wake word: **"hey windows"**. You must say this first, followed by your command in the same sentence.

Here are some examples of commands you can use:

*   **Open an Application:** The application automatically finds most of your installed programs. Just say the name of the program you want to open.
    *   `"hey windows, open notepad"`
    *   `"hey windows, open chrome"`
    *   `"hey windows, open spotify"`

*   **Control the Volume:**
    *   `"hey windows, volume up"` (Increases volume by 10%)
    *   `"hey windows, volume down"` (Decreases volume by 10%)
    *   `"hey windows, mute"` (Toggles mute)

*   **Control Your Computer:**
    *   `"hey windows, sleep"`
    *   `"hey windows, shutdown"`
    *   `"hey windows, restart"`

**Important:** When you tell the system to shut down or restart, a confirmation box will pop up to make sure you don't do it by accident. The "sleep" command does not require confirmation.

---

## For Developers

This section provides instructions for setting up the project, running tests, and building the application from the source code.

### Project Setup

1.  **Prerequisites:** Make sure you have Python 3 installed on your system.
2.  **Create a Virtual Environment:** It is recommended to use a virtual environment to manage dependencies.
    ```bash
    python -m venv venv
    ```
3.  **Activate the Environment:**
    ```bash
    .\venv\Scripts\activate
    ```
4.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

### Running the Application from Source

To run the application directly from the Python source code, use the following command:
```bash
python src/main_app.py
```

### Running Tests

The project includes a full suite of automated tests. To run them, use the following command from the root directory of the project:
```bash
PYTHONPATH=. pytest
```

### Building the Executable

To package the application into a standalone `.exe` file, simply run the build script:
```bash
build.bat
```
This script will handle creating a virtual environment, installing dependencies, and running PyInstaller. The final `VoiceControlApp.exe` file will be located in the `dist` directory.

Download the latest application release here https://github.com/JaiRaj22/Voice-control-windows/releases
