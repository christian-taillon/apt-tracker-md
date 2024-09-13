# APT Tracker Markdown Converter

This script converts the APT Groups and Operations Excel file from https://apt.threattracking.com into individual Markdown files.

## Setup

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/apt-tracker-md.git
   cd apt-tracker-md
   ```

2. Create a virtual environment (optional but recommended):
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

3. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

4. Download the "APT Groups and Operations.xlsx" file from https://apt.threattracking.com and place it in the same directory as the script.

## Usage

Run the script with:

```
python apt.py
```

If you want to specify a different Excel file, use the `-f` or `--file` option:

```
python apt.py -f /path/to/your/excel/file.xlsx
```

The script will create directories for each sheet in the Excel file and generate Markdown files for each APT group.

## Note

The script currently creates links using CrowdStrike's naming convention. You can modify the `modify_content` function in the script to use different naming conventions or link to other intelligence source providers.
```
