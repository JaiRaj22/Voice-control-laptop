# Voice Control for windows

Welcome! This is a simple application that lets you control your Windows computer using your voice. You can open programs, change the volume, and even shut down or restart your computer, all by speaking to it.

## For Everyone: How to Get Started

Getting the application running is simple. You don't need to install anything!

1.  **Find the Application:** Look for the `dist` folder in the project files. Inside that, you will find a folder named `VoiceControlApp`.
2.  **Run the Executable:** Open the `dist/VoiceControlApp` folder. Inside, you will see a file named `VoiceControlApp.exe`. Just double-click it to start the application.
3.  **Start Speaking:** A small window will appear with a real-time log. The application is now listening! When you first start it, it will take a moment to scan your computer for installed programs.

### How to Use the Application

The application listens for a special wake word: **"hey system"**. You must say this first, followed by your command in the same sentence.

Here are some examples of commands you can use:

*   **Open an Application:** The application automatically finds most of your installed programs. Just say the name of the program you want to open.
    *   `"hey system, open notepad"`
    *   `"hey system, open chrome"`
    *   `"hey system, open spotify"`

*   **Control the Volume:**
    *   `"hey system, volume up"` (Increases volume by 10%)
    *   `"hey system, volume down"` (Decreases volume by 10%)
    *   `"hey system, mute"` (Toggles mute)

*   **Control Your Computer:**
    *   `"hey system, sleep"`
    *   `"hey system, shutdown"`
    *   `"hey system, restart"`

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
This script will handle creating a virtual environment, installing dependencies, and running PyInstaller. The final executable and all its supporting files will be located in the `dist/VoiceControlApp` directory.
