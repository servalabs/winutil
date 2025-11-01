# Working
Replace Hosts
Windows Cleanup
Activate Windows/Office
Improve Privacy
Winget App Installs

# To Add

# FYI
Winget doesn't play nice with looks and functions, directly calling is always best

### Quick Start

Run this first to allow running local scripts for your user:

```powershell
Set-ExecutionPolicy Unrestricted -Scope CurrentUser
iwr -useb https://raw.githubusercontent.com/servalabs/winutil/main/Run.ps1 | iex
```