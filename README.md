# APT Tracker Markdown Converter

![APT Tracker Demo](apt-tracker-demo.png)

> [!NOTE]
> This script converts the APT Groups and Operations Excel file from https://apt.threattracking.com into individual Markdown files for use in note-taking and intelligence tracking systems like Obsidian.

## Related Project by Ezra Woods

🔍 **[Malpedia to Markdown Converter](https://github.com/shammahwoods/malpedia-to-md)** 
- Created by [Ezra Woods](https://github.com/shammahwoods)
- Converts Malpedia threat intelligence data to Markdown format
- Complimentary tool for threat intelligence documentation

## Prerequisites

> [!WARNING]
> Ensure you have the following installed:
> - Python 3.7+
> - Git

[... rest of the README remains the same ...]

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/christian-taillon/apt-tracker-md.git
   cd apt-tracker-md
   ```

2. Initialize and update submodules:
   ```bash
   git submodule init
   git submodule update
   ```

3. Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
   ```

4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

5. Download the Excel file:
   > [!TIP]
   > Download "APT Groups and Operations.xlsx" from https://apt.threattracking.com and place it in the project directory.

## Usage

Run the script:
```bash
python apt.py
```

### Optional Arguments

- `-f` or `--file`: Specify a custom Excel file
  ```bash
  python apt.py -f /path/to/your/excel/file.xlsx
  ```

## Customization

> [!IMPORTANT]
> You can modify the `modify_content` function in the script to:
> - Use different naming conventions
> - Link to alternative intelligence sources
> - Customize output formatting

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
