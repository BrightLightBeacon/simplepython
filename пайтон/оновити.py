import os
import sys
import shutil
import urllib.request
import urllib.error
import zipfile
import tempfile
import subprocess
from datetime import datetime

# GitHub repository details
REPO_URL = "https://github.com/BrightLightBeacon/simplepython"
ZIP_URL = f"{REPO_URL}/archive/refs/heads/main.zip"
API_ZIP_URL = "https://api.github.com/repos/BrightLightBeacon/simplepython/zipball/main"

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
            
        print("Git repository detected. Fetching changes...")
        # Fetch remote changes
        fetch_result = subprocess.run(["git", "fetch", "origin", "main"], cwd=root_dir)
        if fetch_result.returncode != 0:
            print("Git fetch failed. Trying ZIP fallback...")
            return False
            
        # Parse diff between HEAD and origin/main to show detailed logs
        # We pass '-c core.quotepath=false' to output Cyrillic characters correctly as UTF-8
        diff_proc = subprocess.run(
            ["git", "-c", "core.quotepath=false", "diff", "--name-status", "HEAD", "origin/main"],
            cwd=root_dir, capture_output=True
        )
        
        diff_lines = diff_proc.stdout.decode('utf-8', errors='ignore').strip().split('\n')
        to_update = []
        to_skip = []
        to_delete = []
        has_changes = False
        
        for line in diff_lines:
            if not line.strip():
                continue
            parts = line.split('\t')
            if len(parts) >= 2:
                status, file_path = parts[0], parts[1]
                if should_update_file(file_path):
                    if status.startswith('D'):
                        to_delete.append(file_path)
                    elif status.startswith('A'):
                        to_update.append((file_path, "ADDED"))
                    else:
                        to_update.append((file_path, "UPDATED"))
                    has_changes = True
                else:
                    to_skip.append(file_path)
                    
        if not has_changes and not to_skip:
            print("No updates available.")
            return True
            
        print("\nChanges to apply:")
        for path, action in to_update:
            print(f"  [{action}] {path}")
        for path in to_delete:
            print(f"  [DELETED] {path}")
        for path in to_skip:
            print(f"  [SKIPPED] {path} (Preserving local copy)")
            
        if not to_update and not to_delete:
            print("\nNo allowed files need updating.")
            return True
            
        print("\nApplying updates...")
        if to_update:
            files_to_checkout = [path for path, _ in to_update]
            subprocess.run(["git", "checkout", "origin/main", "--"] + files_to_checkout, cwd=root_dir)
            
        for path in to_delete:
            full_path = os.path.join(root_dir, path)
            if os.path.exists(full_path):
                try:
                    os.remove(full_path)
                except Exception:
                    pass
                    
        # Reset staging area to keep git status clean
        subprocess.run(["git", "reset", "HEAD"], cwd=root_dir)
        
        print("\nUpdate completed successfully (Git).")
        return True
    except FileNotFoundError:
        print("Git is not installed or not found in PATH.")
        return False
    except Exception as e:
        print(f"Git update failed: {e}")
        return False

def run_zip_update(root_dir):
    """Download ZIP from GitHub and extract it, backing up replaced files."""
    print("\nTrying update via ZIP archive download...")
    temp_zip = None
    temp_dir = None
    
    token_path = os.path.join(root_dir, ".github_token")
    token = None
    if os.path.exists(token_path):
        try:
            with open(token_path, 'r', encoding='utf-8') as f:
                token = f.read().strip()
        except Exception:
            pass
            
    def download_zip(url, use_token=None):
        headers = {'User-Agent': 'Mozilla/5.0'}
        if use_token:
            headers['Authorization'] = f'Bearer {use_token}'
            headers['Accept'] = 'application/vnd.github+json'
        
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req) as response:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
                shutil.copyfileobj(response, tmp)
                return tmp.name
    
    try:
        # 1. Download the ZIP file
        try:
            if token:
                print("Using saved GitHub token...")
                temp_zip = download_zip(API_ZIP_URL, token)
            else:
                print(f"Downloading update from {ZIP_URL}...")
                temp_zip = download_zip(ZIP_URL)
        except urllib.error.HTTPError as e:
            # 404, 403, or 401 indicates that authentication is required
            if e.code in (401, 403, 404):
                print("\n[Auth Required] This repository is private or requires authentication.")
                if token:
                    print("Saved token appears to be invalid or expired.")
                
                # Prompt user for GitHub Personal Access Token (PAT)
                token = input("Please enter your GitHub Personal Access Token (PAT): ").strip()
                if not token:
                    raise Exception("Authentication token is required to download updates.")
                
                print("Testing token and downloading update...")
                temp_zip = download_zip(API_ZIP_URL, token)
                
                # Save the verified working token
                try:
                    with open(token_path, 'w', encoding='utf-8') as f:
                        f.write(token)
                    print(f"Saved token to {token_path} (added to .gitignore).")
                except Exception as save_err:
                    print(f"Warning: Could not save token: {save_err}")
            else:
                raise e
        
        # 2. Create a temporary directory to extract into
        temp_dir = tempfile.mkdtemp()
        print("Extracting archive...")
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        # Locate the extracted folder
        contents = os.listdir(temp_dir)
        if len(contents) == 1 and os.path.isdir(os.path.join(temp_dir, contents[0])):
            source_dir = os.path.join(temp_dir, contents[0])
        else:
            source_dir = os.path.join(temp_dir, "simplepython-main")
            if not os.path.exists(source_dir):
                raise Exception("Could not find extracted repository folder.")
        
        # 3. Scan for changes (added, updated, skipped, deleted)
        to_update = []  # list of (norm_rel_path, src_file, dest_file, action)
        to_skip = []    # list of norm_rel_path
        to_delete = []  # list of norm_rel_path
        
        zip_rel_paths = set()
        
        for root, dirs, files in os.walk(source_dir):
            rel_dir = os.path.relpath(root, source_dir)
            for file in files:
                rel_path = os.path.join(rel_dir, file) if rel_dir != "." else file
                norm_rel_path = rel_path.replace('\\', '/')
                zip_rel_paths.add(norm_rel_path)
                
                src_file = os.path.join(root, file)
                dest_file = os.path.join(root_dir, rel_path)
                
                if should_update_file(norm_rel_path):
                    if not os.path.exists(dest_file):
                        to_update.append((norm_rel_path, src_file, dest_file, "ADDED"))
                    else:
                        try:
                            with open(src_file, 'rb') as f1, open(dest_file, 'rb') as f2:
                                if f1.read() != f2.read():
                                    to_update.append((norm_rel_path, src_file, dest_file, "UPDATED"))
                        except Exception:
                            to_update.append((norm_rel_path, src_file, dest_file, "UPDATED"))
                else:
                    if os.path.exists(dest_file):
                        try:
                            with open(src_file, 'rb') as f1, open(dest_file, 'rb') as f2:
                                if f1.read() != f2.read():
                                    to_skip.append(norm_rel_path)
                        except Exception:
                            pass
                            
        # Check for local files in allowed directories that were deleted in remote ZIP
        local_allowed_files = []
        
        # 1. Bat files in root
        for file in os.listdir(root_dir):
            if file.endswith('.bat') and os.path.isfile(os.path.join(root_dir, file)):
                local_allowed_files.append(file)
                
        # 2. Python files in пайтон
        local_payton = os.path.join(root_dir, "пайтон")
        if os.path.exists(local_payton):
            for file in os.listdir(local_payton):
                if file.endswith('.py') and os.path.isfile(os.path.join(local_payton, file)):
                    local_allowed_files.append(os.path.join("пайтон", file).replace('\\', '/'))
                    
        # 3. Any files in шаблони
        local_shablony = os.path.join(root_dir, "шаблони")
        if os.path.exists(local_shablony):
            for root, dirs, files in os.walk(local_shablony):
                rel_dir = os.path.relpath(root, root_dir)
                for file in files:
                    local_allowed_files.append(os.path.join(rel_dir, file).replace('\\', '/'))
                    
        for local_file in local_allowed_files:
            if local_file not in zip_rel_paths:
                if local_file == "пайтон/оновити.py" or local_file == ".github_token":
                    continue
                to_delete.append(local_file)
                
        if not to_update and not to_delete and not to_skip:
            print("No updates available.")
            return True
            
        print("\nChanges to apply:")
        for path, _, _, action in to_update:
            print(f"  [{action}] {path}")
        for path in to_delete:
            print(f"  [DELETED] {path}")
        for path in to_skip:
            print(f"  [SKIPPED] {path} (Preserving local copy)")
            
        if not to_update and not to_delete:
            print("\nNo allowed files need updating.")
            return True
            
        # Create backups of files we are about to overwrite/delete
        backup_dirname = f"update_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        backup_dir = os.path.join(root_dir, backup_dirname)
        has_backups = False
        
        print("\nBacking up modified files...")
        for path, _, dest_file, action in to_update:
            if action == "UPDATED" and os.path.exists(dest_file):
                if not has_backups:
                    os.makedirs(backup_dir, exist_ok=True)
                    has_backups = True
                backup_dest = os.path.join(backup_dir, path)
                os.makedirs(os.path.dirname(backup_dest), exist_ok=True)
                shutil.copy2(dest_file, backup_dest)
                
        for path in to_delete:
            dest_file = os.path.join(root_dir, path)
            if os.path.exists(dest_file):
                if not has_backups:
                    os.makedirs(backup_dir, exist_ok=True)
                    has_backups = True
                backup_dest = os.path.join(backup_dir, path)
                os.makedirs(os.path.dirname(backup_dest), exist_ok=True)
                shutil.copy2(dest_file, backup_dest)
                
        if has_backups:
            print(f"Backup created in: {backup_dirname}")
            
        # Apply updates
        print("Applying updates...")
        for path, src_file, dest_file, action in to_update:
            if path == "пайтон/оновити.py":
                try:
                    shutil.copy2(src_file, dest_file)
                except Exception:
                    print("  [INFO] пайтон/оновити.py is in use. It will be updated on next execution.")
                continue
            os.makedirs(os.path.dirname(dest_file), exist_ok=True)
            shutil.copy2(src_file, dest_file)
            
        for path in to_delete:
            dest_file = os.path.join(root_dir, path)
            if os.path.exists(dest_file):
                try:
                    os.remove(dest_file)
                except Exception:
                    pass
                    
        print("\nUpdate completed successfully (ZIP).")
        return True
        
    except Exception as e:
        print(f"\n[ERROR] ZIP update failed: {e}")
        if token_path and os.path.exists(token_path):
            try:
                os.remove(token_path)
            except Exception:
                pass
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
    
    success = run_git_update(root_dir)
    if not success:
        success = run_zip_update(root_dir)
        
    if not success:
        print("\n[ERROR] Update failed.")
        
    input("\nНатисніть Enter для виходу / Press Enter to exit...")
    if not success:
        sys.exit(1)
