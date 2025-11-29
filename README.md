# Software-Quality-Analysis
Analyzing Software Quality and Maintainability in Object-Oriented Systems using Software Metrics.
Effective evaluation of software quality and maintainability is compulsory for successful object-oriented system
development, and the potential of software metrics in achieving these goals are investigated in this research. To evaluate the quality of software, this research employs software metrics to identify potential errors and weaknesses in object-oriented systems. 
# This Python project downloads a public GitHub repository, analyzes its commit history and Python source files, and generates Excel reports with detailed metrics.

## Features

- **Download & Unzip** any public GitHub repository using a personal access token.
- **Analyze Commits:** Lists all committers, their emails, commit counts, dates, and messages.
- **Python Code Metrics:** For each Python file, extracts comments, classes, methods, comment percentage, ATFD, WMC, and RFC metrics.
- **Excel Reports:** Generates two Excel files:
  - `repository_analysis.xlsx` (commit and contributor info)
  - `python_file_analysis.xlsx` (code metrics for each Python file)

## Usage
1. Generate a github token

2. Install requirements:
    ```
    pip install requests openpyxl pandas others
    ```
3. Run the script:
    ```
    python python_script.py
    ```
4. Enter the GitHub repository URL and your personal access token when prompted.

## Output

- `repository_analysis.xlsx`
- `python_file_analysis.xlsx`

