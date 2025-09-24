# Email Automation Desktop - Troubleshooting Guide

## Error: "No module named 'ast'"

### Problem
Error muncul saat menjalankan aplikasi yang sudah di-build, menunjukkan modul `ast` tidak ditemukan.

### Root Cause
PyInstaller tidak menyertakan modul `ast` dan beberapa modul built-in Python lainnya yang diperlukan aplikasi.

### Solution Applied
1. **Fixed spec file** - Added critical built-in modules to `hiddenimports`:
```python
hiddenimports=[
    # Python built-in modules (critical)
    'ast',
    'inspect', 
    'copy',
    'pickle',
    'struct',
    'operator',
    'weakref',
    'gc',
    'io',
    'codecs',
    'encodings',
    'encodings.utf_8',
    'encodings.cp1252',
    'locale',
    'warnings',
    'linecache',
    'keyword',
    'token',
    'tokenize',
    # ... other modules
]
```

2. **Removed from excludes** - Removed `ast` from the exclusion list

## How to Fix and Test

### Step 1: Clean Rebuild
```powershell
# Clean rebuild with fixed spec
.\build.ps1 -Clean -InstallDeps
```

### Step 2: Debug Build (if issues persist)
```powershell
# Build with debug console for troubleshooting
.\debug_build.ps1 -Clean -InstallDeps
```

### Step 3: Test Execution
```powershell
# Test the built executable
dist\EmailAutomation.exe

# Or test directory distribution (usually more stable)
dist\EmailAutomation_dist\EmailAutomation.exe
```

## Common Build Issues & Solutions

### Issue 1: Import Errors
**Symptoms**: Missing module errors (`No module named 'X'`)

**Solutions**:
1. Add missing modules to `hiddenimports` in spec file
2. Verify module is installed in build environment
3. Check if module is accidentally excluded

**Debug Steps**:
```powershell
# Test imports before building
.\debug_build.ps1 -InstallDeps
```

### Issue 2: Application Won't Start
**Symptoms**: Application closes immediately, no window appears

**Solutions**:
1. Build with debug console enabled
2. Check for missing data files or resources
3. Verify all dependencies are properly packaged

**Debug Steps**:
```powershell
# Build with console output
.\debug_build.ps1 -Clean -InstallDeps

# Run and check console output
dist\EmailAutomation.exe
```

### Issue 3: GUI Issues
**Symptoms**: Blank window, missing icons, styling problems

**Solutions**:
1. Verify qfluentwidgets resources are included
2. Check theme files and stylesheets
3. Ensure PySide6 modules are properly packaged

**Debug Steps**:
1. Check if `qfluentwidgets` resources are in build
2. Test with directory distribution instead of single-file

### Issue 4: File Access Errors
**Symptoms**: Cannot read/write config files, templates, or database

**Solutions**:
1. Check if data directories are included in build
2. Verify write permissions in application directory
3. Use absolute paths for critical files

**Debug Steps**:
1. Check `dist/` folder for required directories
2. Test file operations in development environment first

## Build Environment Requirements

### Python Version
- **Minimum**: Python 3.9+
- **Recommended**: Python 3.11+

### Dependencies Check
```powershell
# Verify all dependencies are installed
python -c "
import ast, PySide6.QtCore, qfluentwidgets, watchdog, jinja2, win32com.client
print('All critical modules available')
"
```

### Virtual Environment (Recommended)
```powershell
# Create and activate virtual environment
python -m venv venv
.\venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements_optimized.txt

# Build
.\build.ps1 -InstallDeps
```

## Testing Checklist

### Before Building
- [ ] Python 3.9+ installed
- [ ] All dependencies installed (`pip list`)
- [ ] Virtual environment activated (recommended)
- [ ] No import errors when testing modules

### After Building  
- [ ] Executable files created in `dist/` folder
- [ ] Application starts without errors
- [ ] Main window displays properly
- [ ] All navigation tabs work
- [ ] Configuration can be loaded/saved
- [ ] Database connection works
- [ ] Templates can be loaded

### Runtime Testing
- [ ] File monitoring can be started/stopped
- [ ] Email sending works (test email)
- [ ] Logs are generated properly
- [ ] Application closes cleanly

## Debug Tools

### Debug Build Script
Use `debug_build.ps1` for troubleshooting:
```powershell
# Full debug build with console output
.\debug_build.ps1 -Clean -InstallDeps
```

Features:
- Console output enabled for debugging
- Import testing before build
- UPX compression disabled for faster debugging
- Automatic application testing

### Manual Debug Steps
1. **Test imports**:
```python
python -c "
try:
    import ast
    print('✓ ast module available')
except ImportError as e:
    print('✗ ast module missing:', e)
"
```

2. **Check PyInstaller hooks**:
```powershell
python -m PyInstaller --help
python -c "import PyInstaller; print(PyInstaller.__version__)"
```

3. **Verbose build**:
```powershell
python -m PyInstaller --noconfirm --clean --debug=all email_automation_optimized.spec
```

## Performance Optimization

### Reduce Build Size
1. **Remove unused dependencies** from requirements
2. **Add more exclusions** to spec file (carefully)
3. **Use UPX compression** (enabled by default)

### Faster Startup
1. **Use directory distribution** instead of single-file
2. **Disable UPX** for development builds
3. **Pre-compile Python modules**

### Memory Usage
1. **Limit concurrent operations**
2. **Use efficient data structures**
3. **Proper resource cleanup**

## Getting Help

If issues persist after following this guide:

1. **Check console output** from debug build
2. **Test individual components** separately
3. **Verify environment setup** matches requirements
4. **Try directory distribution** instead of single-file
5. **Review PyInstaller logs** for specific errors

### Common Error Patterns

| Error Pattern | Likely Cause | Solution |
|---------------|--------------|----------|
| `No module named 'X'` | Missing hiddenimport | Add to spec file |
| `Failed to execute script` | Missing built-in module | Check excludes list |
| Blank/frozen window | GUI resource missing | Verify qfluentwidgets resources |
| File not found errors | Missing data files | Check datas section in spec |
| Permission denied | Write access issues | Run as admin or check paths |

---

**Remember**: Always test in a clean environment that matches your deployment target!