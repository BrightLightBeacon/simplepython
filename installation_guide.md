# Installation & Update Setup Guide for Private Repository

This guide explains how to install the project files on a target PC and configure the one-click update feature (`оновити.bat`) to work with your private GitHub repository (`BrightLightBeacon/simplepython`).

Choose the method that best matches your target user's technical background:

---

## Method A: ZIP-only Deployment (No Git Required)
*Recommended for non-technical users. It does not require installing Git on their PC.*

### 1. Distribute the Files
1. Create a ZIP archive of your local folder. **Do not include** the `.git/` folder or any backup folders.
2. Send this ZIP file to the target user (via email, shared drive, etc.).
3. The user extracts the ZIP file into any folder on their PC (e.g., `C:\simplepython`).

### 2. Generate an Access Token (PAT)
Since the repository is private, the target user needs a GitHub Personal Access Token (PAT) with read access to download updates:
1. Log into the GitHub account that has access to the repository.
2. Go to **Settings** -> **Developer settings** -> **Personal access tokens** -> **Tokens (classic)**.
3. Click **Generate new token (classic)**.
4. Set a name (e.g., "target-pc-updater"), choose an expiration date, and check the **`repo`** scope (or **`public_repo`** if the repository is public but write-protected).
5. Click **Generate token** and copy the generated token string.

### 3. First-Time Authentication on Target PC
1. On the target PC, open the project folder and double-click **`оновити.bat`**.
2. Because the repository is private and no Git is installed, the console will prompt:
   ```text
   [Auth Required] This repository is private or requires authentication.
   Please enter your GitHub Personal Access Token (PAT): 
   ```
3. Paste the token and press **Enter**.
4. The script verifies the token, downloads the latest ZIP archive via the GitHub API, applies the updates, and saves the token to a local, git-ignored file named `.github_token`.
5. **Subsequent updates** will run silently and automatically without prompting again.

---

## Method B: Git-based Deployment (Git Required)
*Recommended for developers or users who already have Git installed.*

### 1. Clone the Repository
On the target PC, open the command line and run:
```cmd
git clone https://github.com/BrightLightBeacon/simplepython.git C:\simplepython
```
*(Or use the SSH URL `git@github.com:BrightLightBeacon/simplepython.git` if they prefer SSH).*

### 2. Configure Authentication
* **If cloned via HTTPS**: The first time the user runs `оновити.bat`, the Windows **Git Credential Manager** pop-up will appear. The user signs into their GitHub account (or enters a PAT). Windows caches these credentials automatically, allowing future updates to run with one click.
* **If cloned via SSH**: The user generates an SSH keypair on their machine (`ssh-keygen`) and registers their public key in their GitHub account settings (**SSH and GPG keys**). The update script will use the local SSH agent to authenticate automatically.

### 3. Run Updates
The user simply double-clicks **`оновити.bat`** to fetch and apply updates.
