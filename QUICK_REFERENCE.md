# Quick Reference - Windows EXE Build

## 🚀 One-Time Setup

```bash
cd /Users/kunal/Canvex/FinalBuildMac

# Initialize git repo
git init
git add .
git commit -m "Add Windows build automation"

# Set GitHub remote (replace URL with your repo)
git remote add origin https://github.com/YOUR_USERNAME/Canvex.git
git branch -M main
git push -u origin main
```

## 📊 After Setup - Every Time You Update

```bash
# Make your changes to Canvex.py
nano Canvex.py

# Commit and push
git add .
git commit -m "Your change description"
git push
```

**Result:** GitHub automatically builds Windows .exe in 2-3 minutes!

## 📥 Download Your Windows EXE

1. Go to: `https://github.com/YOUR_USERNAME/Canvex/actions`
2. Click latest **"Build Windows EXE"** workflow
3. Download **Canvex-Windows** artifact
4. Extract `Canvex.exe`

## 🏷️ Create GitHub Release

```bash
# Tag and push to create a formal release
git tag v1.0.0
git push origin v1.0.0
```

**Result:** GitHub creates Release page with .exe attached!

## 📁 Files Created

```
✅ .github/workflows/build-windows.yml  (Workflow config)
✅ requirements.txt                     (Python packages)
✅ README.md                            (Documentation)
✅ GITHUB_ACTIONS_SETUP.md              (Detailed guide)
✅ WINDOWS_BUILD_SUMMARY.md             (This setup summary)
```

## 🔧 If Build Fails

1. Check GitHub Actions logs
2. Common fixes:
   - Verify `requirements.txt` has all packages
   - Check `Canvex.spec` exists
   - Ensure `Canvex.py` has no syntax errors

## 💾 What Gets Built

| Platform | Output | Source |
|----------|--------|--------|
| macOS | `dist/Canvex.app` | Your Mac |
| Windows | `dist/Canvex.exe` | GitHub Actions |

## 📞 Support

- **Setup Help**: See `GITHUB_ACTIONS_SETUP.md`
- **Usage Help**: See `README.md`
- **Troubleshooting**: Check GitHub Actions workflow logs

---

**That's it!** You now have automated Windows builds. 🎉
