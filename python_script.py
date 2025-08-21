import requests  # Import the requests library for making HTTP requests
import os  # Import the os library for file and directory operations
from zipfile import ZipFile  # Import ZipFile for working with zip files
import openpyxl  # Import openpyxl for working with Excel files
import re  # Import re for regular expressions
import ast  # Import ast for abstract syntax tree operations
from collections import Counter  # Import Counter for counting elements in lists
import pandas as pd  # Import pandas for data manipulation

# Define a class to calculate RFC (Response For a Class) metric
class RFCCalculator(ast.NodeVisitor):
    def __init__(self):
        self.class_rfc = {}
        

    def visit_ClassDef(self, node):
        class_name = node.name
        method_info = self.collect_method_info(node.body)
        self.class_rfc[class_name] = method_info

    def collect_method_info(self, body):
        method_info = {}

        for subnode in body:
            if isinstance(subnode, ast.FunctionDef):
                method_name = subnode.name
                method_calls = self.find_method_calls(subnode)
                method_info[method_name] = {
                    "method_calls": method_calls,
                    "method_count": len(method_calls)
                }

        return method_info

    def find_method_calls(self, node):
        method_calls = set()

        for subnode in ast.walk(node):
            if isinstance(subnode, ast.Call) and isinstance(subnode.func, ast.Name):
                method_calls.add(subnode.func.id)

        return method_calls

# Function to download and unzip a GitHub repository
def download_and_unzip_github_repository(repo_url, access_token):
    # Extract username and repository name from the URL
    _, _, _, username, repository = repo_url.rstrip('/').split('/')
    zip_file_name = f"{username}_{repository}_master.zip"  # Name of the zip file
    api_url = f"https://api.github.com/repos/{username}/{repository}/zipball/master"  # GitHub API URL for downloading the repo
    headers = {'Authorization': f'token {access_token}'}  # Authorization header with access token
    response = requests.get(api_url, headers=headers)  # Make a GET request to download the repo

    if response.status_code == 200:
        # Write the content of the response to a zip file
        with open(zip_file_name, 'wb') as zip_file:
            zip_file.write(response.content)
        print(f"Repository downloaded successfully as {zip_file_name}")

        # Extract the content of the zip file
        with ZipFile(zip_file_name, 'r') as zip_ref:
            zip_ref.extractall()
        print(f"Repository unzipped successfully.")

        # Analyze and create Excel files
        analyze_and_create_excel(repo_url, access_token)
        analyze_python_files_and_create_excel()
    else:
        print(f"Failed to download repository. Status code: {response.status_code}")

# Function to analyze the repository and create an Excel file
def analyze_and_create_excel(repo_url, access_token):
    workbook = openpyxl.Workbook()  # Create a new Excel workbook
    sheet = workbook.active  # Get the active sheet
    sheet.append(["Repository Information"])  # Add header for repository information
    sheet.append(["Repo Name", "All Commitors", "Total Number of Commits"])  # Add columns

    print("\nAnalyzing the unzipped repository:")

    # Extract owner and repo name from the URL
    _, _, _, owner, repo = repo_url.rstrip('/').split('/')
    api_url = 'https://api.github.com/'  # Base GitHub API URL
    headers = {'Authorization': f'token {access_token}'}  # Authorization header with access token
    commits_url = f'{api_url}repos/{owner}/{repo}/commits'  # GitHub API URL for commits
    response = requests.get(commits_url, headers=headers)  # Make a GET request to fetch commits

    if response.status_code == 200:
        commits = response.json()  # Parse the JSON response
        repo_name = repo  # Repository name
        commitors = set()  # Set to store unique committers
        total_commits = len(commits)  # Total number of commits

        commit_counts = {}  # Dictionary to count commits per developer

        for commit in commits:
            developer_name = commit['commit']['author']['name']
            developer_email = commit['commit']['author']['email']

            if developer_name in commit_counts:
                commit_counts[developer_name] += 1
            else:
                commit_counts[developer_name] = 1

            commitors.add(developer_name)

        # Add repository information to the sheet
        sheet.append([repo_name, ", ".join(commitors), total_commits])
        sheet.append([])  # Empty row as a separator
        sheet.append(["Developers Information"])
        sheet.append(["Commiter's Name", "Committer's Email", "Number of Commits", "Commit Date and Time", "Commit Message"])

        # Add commit information to the sheet
        for commit in commits:
            developer_name = commit['commit']['author']['name']
            developer_email = commit['commit']['author']['email']
            commit_date = commit['commit']['author']['date']
            commit_message = commit['commit']['message']

            sheet.append([developer_name, developer_email, commit_counts[developer_name], commit_date, commit_message])

        excel_file_name = "repository_analysis.xlsx"  # Name of the Excel file
        workbook.save(excel_file_name)  # Save the workbook

        print(f"\nExcel sheet created successfully: {excel_file_name}")

        for developer, count in commit_counts.items():
            print(f"{developer} has {count} commits.")
    else:
        print(f"Error fetching commits: {response.status_code}")

# Function to analyze Python files and create an Excel file
def analyze_python_files_and_create_excel():
    print("\nAnalyzing Python files:")

    # Find all Python files in the current directory and subdirectories
    python_files = [os.path.join(root, file) for root, dirs, files in os.walk(".") for file in files if file.endswith(".py")]

    workbook_python = openpyxl.Workbook()  # Create a new Excel workbook
    sheet_python = workbook_python.active  # Get the active sheet
    sheet_python.append(["Python File Analysis"])  # Add header for Python file analysis
    sheet_python.append(["File Name", "Comments", "Total Comments", "Class Names", "Method Names", "Comment Percentage", "Total Lines of Code", "ATFD", "Total WMC", "Methods (RFC)"])  # Add columns

    for file_path in python_files:
        print(f"Analyzing Python file: {file_path}")
        file_name = os.path.basename(file_path)  # Get the file name

        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.readlines()  # Read the content of the file

            # Extract comments, methods, classes, and calculate metrics
            comments, comment_lines, methods, class_names, comment_percentage, total_lines = extract_comments_methods(content)
            method_names = [method[0] for method in methods]
            class_names_list = ', '.join(class_names).split(', ')
            method_counter = Counter(method_names)
            methods_used_in_multiple_classes = [method for method, count in method_counter.items() if count > 1]
            atfd = calculate_atfd(content)
            wmc = calculate_wmc(content)
            rfc = calculate_rfc(content)

            # Add Python file analysis information to the sheet
            sheet_python.append([file_name, '\n'.join(comments), comment_lines, ', '.join(class_names_list), ', '.join(method_names), comment_percentage, total_lines, atfd, wmc, rfc])
        except Exception as e:
            print(f"Error analyzing Python file {file_path}: {e}")

    excel_file_name_python = "python_file_analysis.xlsx"  # Name of the Excel file
    workbook_python.save(excel_file_name_python)  # Save the workbook

    print(f"\nExcel sheet for Python file analysis created successfully: {excel_file_name_python}")

# Function to extract comments, methods, classes, and calculate comment percentage
def extract_comments_methods(content):
    comment_pattern = r'#.*'  # Pattern for comments
    method_pattern = r'def\s+(\w+)\s*\((.*?)\):'  # Pattern for method definitions
    class_pattern = r'class\s+(\w+)\s*:'  # Pattern for class definitions

    comments = []  # List to store comments
    for line in content:
        if not line.strip():
            continue
        match = re.match(comment_pattern, line.strip())
        if match:
            comments.append(match.group())

    methods = re.findall(method_pattern, ''.join(content))
    class_names = re.findall(class_pattern, ''.join(content))
    total_lines = len([line for line in content if line.strip()])
    comment_lines = len(comments)
    comment_percentage = (comment_lines / total_lines) * 100

    return comments, comment_lines, methods, class_names, comment_percentage, total_lines

# Function to calculate ATFD (Access To Foreign Data)
def calculate_atfd(content):
    tree = ast.parse(''.join(content))  # Parse the content into an AST
    atfd = 0
    foreign_attributes = set()

    for method_node in ast.walk(tree):
        if isinstance(method_node, ast.FunctionDef):
            for subnode in ast.walk(method_node):
                if isinstance(subnode, ast.Attribute):
                    attr_owner = subnode.value.id if isinstance(subnode.value, ast.Name) else None
                    if attr_owner and attr_owner != method_node.name:
                        foreign_attributes.add(subnode.attr)

    atfd = len(foreign_attributes)
    return atfd

# Function to calculate WMC (Weighted Methods per Class)
def calculate_wmc(content):
    tree = ast.parse(''.join(content))  # Parse the content into an AST
    wmc = sum(1 for method_node in ast.walk(tree) if isinstance(method_node, ast.FunctionDef))  # Count the number of methods
    return wmc

# Function to calculate RFC (Response For a Class)
def calculate_rfc(content):
    tree = ast.parse(''.join(content))  # Parse the content into an AST
    rfc_calculator = RFCCalculator()  # Instantiate the RFCCalculator
    rfc_calculator.visit(tree)  # Visit the AST nodes

    rfc_info = rfc_calculator.class_rfc
    total_rfc = sum(sum(info['method_count'] for info in method_info.values()) for method_info in rfc_info.values())

    return total_rfc

if __name__ == "__main__":
    github_url = input("Enter the GitHub repository URL: ")  # Prompt the user for the GitHub repository URL
    access_token = input("Enter your GitHub access token: ")  # Prompt the user for the GitHub access token

    download_and_unzip_github_repository(github_url, access_token)  # Call the function to download and unzip the repository
