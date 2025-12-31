# GitHub Actions Setup Guide

Follow these steps to set up automatic Windows .exe builds:

## Step 1: Initialize Git Repository

If you haven't already, initialize git and push to GitHub:

```bash
cd /Users/kunal/Canvex/FinalBuildMac

# Initialize git (if not already done)
git init
git add .
git commit -m "Initial commit: Add Canvex app with Windows build workflow"

# Add remote (replace YOUR_USERNAME and YOUR_REPO)
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git

# Create main branch and push
git branch -M main
git push -u origin main
```

## Step 2: GitHub Repository Setup

1. Go to https://github.com/new
2. Create a new repository called `Canvex` (or your preferred name)
3. Make sure it's **Public** (required for free GitHub Actions)
4. Copy the repository URL

## Step 3: Push Your Code

Use the git commands from Step 1 to push your code to GitHub.

## Step 4: Verify Workflow

1. Go to your GitHub repository
2. Click on the **Actions** tab
3. You should see "Build Windows EXE" workflow listed
4. Any push to `main` branch will trigger an automatic build

## Step 5: Access Built EXE

After a successful build:

1. Go to **Actions** tab
2. Click on the latest "Build Windows EXE" workflow run
3. Scroll down to **Artifacts** section
4. Download `Canvex-Windows` (contains `Canvex.exe`)

## Step 6: Create Release Builds

To create a proper GitHub Release with the .exe:

```bash
# Tag your commit
git tag v1.0.0
git push origin v1.0.0
```

This will:
1. Trigger the Windows build
2. Automatically create a GitHub Release
3. Attach the `.exe` to the release

## Workflow Features

- **Triggers**: Every push to `main` or `master` branch
- **Manual Trigger**: Can be triggered manually via Actions tab
- **Artifacts**: Available for 30 days
- **Release**: Create tagged releases for official versions

## Troubleshooting

### Build Fails with "requirements not found"

Make sure `requirements.txt` is in the root directory with all dependencies listed.

### EXE runs but crashes on Windows

1. Check the GitHub Actions logs for errors
2. Common issues:
   - Missing dependencies in `requirements.txt`
   - Missing WebDriver for Selenium
   - File path issues (use relative paths)

### Workflow doesn't trigger

1. Verify `.github/workflows/build-windows.yml` exists
2. Check that you're pushing to `main` or `master` branch
3. Go to Actions tab and check for any error messages

## File Structure

```
FinalBuildMac/
├── Canvex.py
├── Canvex.spec
├── requirements.txt              ← Must include all dependencies
├── .github/
│   └── workflows/
│       └── build-windows.yml     ← Workflow file
├── README.md
└── GITHUB_ACTIONS_SETUP.md       ← This file
```

## Example Workflow

1. **Day 1**: Push code → Automatic Windows build created
2. **Day 2**: Make improvements → Push update → New Windows build
3. **Day 10**: Ready for release → Tag v1.0.0 → Release created with .exe

## Need Help?

Check the GitHub Actions logs:
1. Go to Actions tab
2. Click on failed workflow
3. Click on "build-windows" job
4. Expand "Build Windows EXE" step to see error details
