# Synchro-Automation

An automation tool for the Synchro application.

## Overview

The `Synchro-Automation` tool automates interactions with the Synchro traffic analysis software. It provides functionalities to synchronize data, import/export data to/from Synchro, convert model volumes to Synchro UTDF, generate reports, and more. This tool utilizes Python and various libraries like `tkinter` for the GUI, `pyautogui` for automation, and `openpyxl` for handling Excel files.

## Features

- **GUI for Easy Interaction**: The tool provides a graphical user interface for selecting model files, Synchro folders, and configuring settings.
- **Automation with Synchro**: Automates interactions with the Synchro application using `pyautogui`.
- **Data Synchronization**: Synchronizes data between the model files and the Synchro application.
- **Import/Export Functionality**: Facilitates importing and exporting data to/from Synchro.
- **Report Generation**: Generates reports based on the synchronized data.

## Installation

### Prerequisites

- Python (version 3.6 or above)
- Synchro software

### Steps

1. Clone the repository:
    ```bash
    git clone https://github.com/deringar/Synchro-Automation.git
    cd Synchro-Automation
    ```

2. Install the required packages:
    ```bash
    pip install -r requirements.txt
    ```

3. Configure the settings:
    - Open the `settings.json` file and configure it according to your environment.
    - Specify paths and other necessary settings for the Synchro application.

## Usage

1. Start the GUI application:
    ```bash
    python main.py
    ```

2. Use the GUI to:
    - Select the model file location.
    - Select the Synchro file folder.
    - Configure additional settings.

3. Start the synchronization process by clicking the "Start" button in the GUI.

## Main Components

### `gui.py`

- **MainWindow**: Represents the main window of the application. Contains methods to set up the UI, create various UI elements (labels, buttons), and handle user interactions. Provides functionality to launch other components like settings and file matching tools.

### `main.py`

- The main entry point for the application. Initializes the `Base` and `MainWindow` classes and starts the application.

## Contributing

1. Fork the repository.
2. Create a new feature branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -m 'Add some feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a Pull Request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Authors: Phillip Gotthelf, Alex Dering

For more details, refer to the [documentation](https://github.com/deringar/Synchro-Automation/blob/main/README.md).