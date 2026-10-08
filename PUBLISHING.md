# Publishing to GitHub

## Prepare

1. Review README.md and the example files.
2. Keep real API keys, proxy credentials, personal spreadsheets, browser sessions, screenshots, and local configuration on your computer.
3. In GitHub Settings → Emails, enable email privacy and copy your GitHub-provided noreply email. Git author names and emails appear in commits, including public commits.
4. Set your repository-specific identity using the commands below. Replace the placeholders.

~~~powershell
git config user.name "YOUR_GITHUB_USERNAME"
git config user.email "YOUR_GITHUB_NOREPLY_EMAIL"
~~~

## Review and commit

If this folder has not been initialized as a repository, run git init -b main first.

~~~powershell
git add .gitignore README.md PUBLISHING.md package.json package-lock.json clickbot.js ads-config.example.json groups.example.json proxies.example.txt
git status --short
git diff --cached --stat
git diff --cached
~~~

Only the nine listed files should be staged. Check the actual diff for private data, then commit:

~~~powershell
git commit -m "Prepare browser automation project for GitHub"
~~~

## Create the GitHub repository

Create an empty repository named proxypilot. Start private while reviewing. Do not add a README, .gitignore, or license in GitHub's creation screen because your local project already has its own files.

Copy the HTTPS repository URL and replace the placeholder below:

~~~powershell
git remote add origin https://github.com/YOUR_USERNAME/proxypilot.git
git push -u origin main
~~~

Authenticate through your Git credential manager or GitHub's supported sign-in flow. Never put an access token in a remote URL. If origin already exists, inspect git remote -v before changing it.

After upload, inspect the Files and Commits views. When satisfied, you can change repository visibility to public. Add a concise description, relevant topics such as playwright, nodejs, adspower, and browser-automation, and choose a license if you want to grant reuse permissions.

## Do and do not

- Do commit the lockfile and credential-free examples.
- Do review new files and add their paths to the .gitignore allowlist when appropriate.
- Do rotate any credential that has previously been exposed.
- Do not upload the entire folder through the website or as a ZIP; that bypasses Git's ignore rules.
- Do not use git add -f for private files.
- Do not share real cookies, session dumps, proxy accounts, API keys, spreadsheets, or screenshots containing account details.
- Do not assume a private repository makes committing secrets safe.
- Do not assume deleting a secret from the latest version removes it from old commits. Revoke exposed credentials and clean affected history before publishing.

Official guide: https://docs.github.com/en/migrations/importing-source-code/using-the-command-line-to-import-source-code/adding-locally-hosted-code-to-github
