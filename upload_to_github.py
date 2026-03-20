#!/usr/bin/env python3
"""
Upload Python project to GitHub using GitHub REST API
"""

import os
import json
import base64
import requests
from pathlib import Path

# Configuration
OWNER = "ishu128"
REPO = "Accounting-Month-End-Automation-RPA-"
BRANCH = "main"

# Get token from environment or console
TOKEN = os.environ.get("GITHUB_TOKEN")
if not TOKEN:
    TOKEN = input("Enter your GitHub Personal Access Token: ")

# GitHub API URLs
API_BASE = "https://api.github.com"
REPO_URL = f"{API_BASE}/repos/{OWNER}/{REPO}"

# Files and paths to ignore
IGNORE_PATTERNS = {
    "__pycache__",
    ".pyc",
    ".git",
    "upload_to_github.py",
    ".gitignore"
}

def should_ignore(path):
    """Check if path should be ignored"""
    path_str = str(path)
    return any(pattern in path_str for pattern in IGNORE_PATTERNS)

def get_all_files(directory):
    """Recursively get all files to upload"""
    files = []
    for root, dirs, filenames in os.walk(directory):
        # Filter directories
        dirs[:] = [d for d in dirs if not should_ignore(d)]
        
        for filename in filenames:
            filepath = os.path.join(root, filename)
            if not should_ignore(filepath):
                files.append(filepath)
    return files

def get_file_content(filepath):
    """Read file content"""
    with open(filepath, "rb") as f:
        return f.read()

def upload_file(filepath, relative_path):
    """Upload a single file to GitHub"""
    headers = {
        "Authorization": f"token {TOKEN}",
        "Accept": "application/vnd.github.v3+json"
    }
    
    content = get_file_content(filepath)
    
    # Encode content to base64
    encoded_content = base64.b64encode(content).decode()
    
    # Prepare request data
    url = f"{REPO_URL}/contents/{relative_path}"
    
    data = {
        "message": f"Add {relative_path}",
        "content": encoded_content,
        "branch": BRANCH
    }
    
    try:
        # Check if file exists
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            # File exists, update it
            file_sha = response.json()["sha"]
            data["sha"] = file_sha
            print(f"Updating: {relative_path}")
        else:
            print(f"Creating: {relative_path}")
        
        # Upload/update file
        response = requests.put(url, headers=headers, json=data)
        
        if response.status_code in [201, 200]:
            print(f"  ✓ Success")
            return True
        else:
            print(f"  ✗ Error: {response.status_code}")
            print(f"  Response: {response.text}")
            return False
    except Exception as e:
        print(f"  ✗ Exception: {e}")
        return False

def main():
    """Main upload function"""
    print(f"Uploading to {OWNER}/{REPO}...")
    print()
    
    # Get project root
    project_root = os.path.dirname(os.path.abspath(__file__))
    
    # Get all files
    files = get_all_files(project_root)
    
    if not files:
        print("No files found to upload!")
        return False
    
    print(f"Found {len(files)} files to upload\n")
    
    # Upload each file
    success_count = 0
    for filepath in files:
        relative_path = os.path.relpath(filepath, project_root)
        relative_path = relative_path.replace("\\", "/")
        
        if upload_file(filepath, relative_path):
            success_count += 1
    
    print()
    print(f"Upload complete: {success_count}/{len(files)} files uploaded successfully")
    
    if success_count == len(files):
        print("✓ All files uploaded successfully!")
        return True
    else:
        print("✗ Some files failed to upload")
        return False

if __name__ == "__main__":
    main()
