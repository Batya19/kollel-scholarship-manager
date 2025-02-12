# Kollel Scholarship Manager

A modern desktop application for managing and calculating Kollel (Jewish seminary) student scholarships based on attendance, punctuality, and academic performance. Built with Python and CustomTkinter.

## Features

- 📊 Automated scholarship calculations based on attendance data
- ⏰ Sophisticated time tracking and punctuality assessment
- 💰 Multiple bonus types support
- 🎯 Configurable attendance targets and thresholds
- 📑 Excel-based input and output
- 🎨 Modern, user-friendly interface

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/batya19/kollel-scholarship-manager.git
   cd kollel-scholarship-manager
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```

2. Click "בחר קובץ לחישוב" to select your Excel input file
3. Enter the number of working days for the period
4. The application will generate "מלגות חודשיות.xlsx" with the calculations

## Input File Format

The input Excel file should contain the following columns:
- זהות (ID)
- שם משפחה (Last Name)
- שם פרטי (First Name)
- כניסה (Entry Time)
- יציאה (Exit Time)
- רצופות (Continuous Study)

## Configuration

Adjust the settings in `config/settings.py` to modify:
- Session times and thresholds
- Base scholarship amounts
- Bonus criteria and amounts

## Building an Executable (Windows)

To generate an `.exe` file for easy distribution, follow these steps:

1. **Ensure dependencies are installed**  
   If you haven't already installed the project dependencies, run:
   ```bash
   pip install -r requirements.txt
   ```

2. **Install `pyinstaller`**  
   ```bash
   pip install pyinstaller
   ```

3. **Build the executable**  
   Run the following command to create an `.exe` file:
   ```bash
   pyinstaller --onefile --windowed main.py
   ```
   - `--onefile`: Packages everything into a single `.exe` file.  
   - `--windowed`: Hides the console window (optional for GUI apps).  

4. **Find the output file**  
   The generated `.exe` will be inside the `dist` folder.  
   You can now distribute `dist/main.exe` to users.


## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0).  
See the [LICENSE](LICENSE) file for details.

## Usage Restrictions

- This software is provided for **personal and educational use only**.  
- **Commercial use is strictly prohibited** without explicit permission.  
- Modifications and redistribution are allowed, but must comply with the Creative Commons Attribution-NonCommercial 4.0 International License.
