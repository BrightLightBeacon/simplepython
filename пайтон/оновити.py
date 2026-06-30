import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
import subprocess
from datetime import datetime

# GitHub repository details
REPO_URL = "https://github.com/BrightLightBeacon/simplepython"
ZIP_URL = f"{REPO_URL}/archive/refs/heads/main.zip"

def should_update_file(rel_path):
    """
    Determines if a file should be updated/added/replaced.
    Only allows:
    1. .bat files in the root or elsewhere.
    2. .py files inside the 'пайтон' folder.
    3. Any files inside the 'шаблони' folder.
    All other files (like 'дебеторка.xlsx' or 'реєстри/*') are ignored.
    """
    # Normalize separators
    norm_path = rel_path.replace('\\', '/')
    parts = norm_path.split('/')
    
    # 1. Any .bat file
    if norm_path.endswith('.bat'):
        return True
        
    # 2. Python files within 'пайтон' folder
    if len(parts) >= 2 and parts[0] == 'пайтон' and norm_path.endswith('.py'):
        return True
        
    # 3. Files within 'шаблони' folder
    if len(parts) >= 2 and parts[0] == 'шаблони':
        return True
        
    return False

def run_git_update(root_dir):
    """Try to update the repository using git pull on specific allowed paths."""
    print("Checking for Git update...")
    try:
        with open(os.devnull, 'w') as devnull:
            git_check = subprocess.run(
                ["git", "rev-parse", "--is-inside-work-tree"],
                cwd=root_dir, stdout=devnull, stderr=devnull
            )
        
        if git_check.returncode != 0:
            print("Not a git repository.")
            return False
            
        print("Git repository detected. Fetching latest changes...")
        # Fetch remote changes
        fetch_result = subprocess.run(["git", "fetch", "origin", "main"], cwd=root_dir)
        if fetch_result.returncode != 0:
            print("Git fetch failed. Will try ZIP download fallback.")
            return False
            
        # Checkout ONLY the allowed folders/files from origin/main
        print("Updating .bat files...")
        subprocess.run(["git", "checkout", "origin/main", "--", "*.bat"], cwd=root_dir)
        
        print("Updating Python scripts...")
        subprocess.run(["git", "checkout", "origin/main", "--", "пайтон/"], cwd=root_dir)
        
        print("Updating templates...")
        subprocess.run(["git", "checkout", "origin/main", "--", "шаблони/"], cwd=root_dir)
        
        # Reset staging area to keep git status clean
        print("Resetting git staging area...")
        subprocess.run(["git", "reset", "HEAD"], cwd=root_dir)
        
        print("\nSuccessfully updated using Git!")
        return True
    except FileNotFoundError:
        print("Git is not installed or not found in PATH.")
        return False
    except Exception as e:
        print(f"Git update failed due to: {e}")
        return False

def run_zip_update(root_dir):
    """Download ZIP from GitHub and extract it, backing up replaced files."""
    print("\nUpdating via ZIP archive download...")
    temp_zip = None
    temp_dir = None
    
    try:
        # 1. Download the ZIP file
        print(f"Downloading update from {ZIP_URL}...")
        headers = {'User-Agent': 'Mozilla/5.0'}
        req = urllib.request.Request(ZIP_URL, headers=headers)
        
        with urllib.request.urlopen(req) as response:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                shutil.copyfileobj(response, tmp)
                temp_zip = tmp.name
        
        # 2. Create a temporary directory to extract into
        temp_dir = tempfile.mkdtemp()
        print("Extracting archive...")
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        # The zip extracts into simplepython-main/ folder
        extracted_folder_name = "simplepython-main"
        source_dir = os.path.join(temp_dir, extracted_folder_name)
        
        if not os.path.exists(source_dir):
            contents = os.listdir(temp_dir)
            if len(contents) == 1 and os.path.isdir(os.path.join(temp_dir, contents[0])):
                source_dir = os.path.join(temp_dir, contents[0])
            else:
                raise Exception("Could not find extracted repository folder.")
        
        # 3. Create a backup of files we are about to overwrite
        backup_dirname = f"оновити_резервна_копія_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        backup_dir = os.path.join(root_dir, backup_dirname)
        has_backups = False
        
        print("Preparing backup for overwritten files...")
        
        for root, dirs, files in os.walk(source_dir):
            rel_path = os.path.relpath(root, source_dir)
            if rel_path == ".":
                target_parent = root_dir
            else:
                target_parent = os.path.join(root_dir, rel_path)
                
            for file in files:
                file_rel_path = os.path.join(rel_path, file) if rel_path != "." else file
                if not should_update_file(file_rel_path):
                    continue
                    
                src_file = os.path.join(root, file)
                dest_file = os.path.join(target_parent, file)
                
                if os.path.exists(dest_file):
                    try:
                        with open(src_file, 'rb') as f1, open(dest_file, 'rb') as f2:
                            if f1.read() != f2.read():
                                if not has_backups:
                                    os.makedirs(backup_dir, exist_ok=True)
                                    has_backups = True
                                
                                backup_dest = os.path.join(backup_dir, rel_path if rel_path != "." else "", file)
                                os.makedirs(os.path.dirname(backup_dest), exist_ok=True)
                                shutil.copy2(dest_file, backup_dest)
                    except Exception:
                        pass
        
        if has_backups:
            print(f"Created backup of modified files in: {backup_dirname}")
            
        # 4. Copy new files over
        print("Applying updates...")
        for root, dirs, files in os.walk(source_dir):
            rel_path = os.path.relpath(root, source_dir)
            if rel_path == ".":
                target_parent = root_dir
            else:
                target_parent = os.path.join(root_dir, rel_path)
                
            for file in files:
                file_rel_path = os.path.join(rel_path, file) if rel_path != "." else file
                if not should_update_file(file_rel_path):
                    continue
                    
                src_file = os.path.join(root, file)
                dest_file = os.path.join(target_parent, file)
                
                # Avoid copying the script itself to prevent file-in-use errors
                if rel_path == "пайтон" and file == "оновити.py":
                    continue
                    
                os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                shutil.copy2(src_file, dest_file)
                
        print("\nSuccessfully updated via ZIP download!")
        return True
        
    except Exception as e:
        print(f"\n[ERROR] ZIP update failed: {e}")
        return False
        
    finally:
        if temp_zip and os.path.exists(temp_zip):
            try:
                os.remove(temp_zip)
            except Exception:
                pass
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass

if __name__ == "__main__":
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(script_dir)
    
    print("=== ОНОВЛЕННЯ ПРОГРАМИ / UPDATE ===")
    print(f"Project directory: {root_dir}\n")
    
    success = run_git_update(root_dir)
    if not success:
        success = run_zip_update(root_dir)
        
    if success:
        print("\n[OK] Update completed successfully!")
    else:
        print("\n[FAILED] Could not update files.")
        sys.exit(1)
