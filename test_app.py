import sys
import os

# Test basic imports
def test_imports():
    print("Testing imports...")

    try:
        import PySide6
        print("[OK] PySide6 available")
    except ImportError:
        print("[FAIL] PySide6 not available")
        return False

    try:
        import qfluentwidgets
        print("[OK] qfluentwidgets available")
    except ImportError:
        print("[FAIL] qfluentwidgets not available")
        return False

    try:
        import watchdog
        print("[OK] watchdog available")
    except ImportError:
        print("[FAIL] watchdog not available")
        return False

    try:
        import jinja2
        print("[OK] jinja2 available")
    except ImportError:
        print("[FAIL] jinja2 not available")
        return False

    print("All imports successful!")
    return True

def test_local_modules():
    print("\nTesting local modules...")

    # Add src to path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    src_dir = os.path.join(current_dir, 'src')
    sys.path.insert(0, src_dir)

    try:
        from core.config_manager import ConfigManager
        print("[OK] ConfigManager imported")
    except Exception as e:
        print(f"[FAIL] ConfigManager failed: {e}")
        return False

    try:
        from core.database_manager import DatabaseManager
        print("[OK] DatabaseManager imported")
    except Exception as e:
        print(f"[FAIL] DatabaseManager failed: {e}")
        return False

    try:
        from utils.logger import setup_logger
        print("[OK] Logger imported")
    except Exception as e:
        print(f"[FAIL] Logger failed: {e}")
        return False

    print("All local modules imported successfully!")
    return True

def test_basic_functionality():
    print("\nTesting basic functionality...")

    try:
        # Add src to path
        current_dir = os.path.dirname(os.path.abspath(__file__))
        src_dir = os.path.join(current_dir, 'src')
        sys.path.insert(0, src_dir)

        from core.config_manager import ConfigManager
        config = ConfigManager()
        print("[OK] ConfigManager created")

        from core.database_manager import DatabaseManager
        db = DatabaseManager(config.get_database_path())
        print("[OK] DatabaseManager created")

        print("Basic functionality test passed!")
        return True

    except Exception as e:
        print(f"[FAIL] Basic functionality test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Email Automation Desktop - Diagnostic Test")
    print("=" * 50)

    # Test imports
    if not test_imports():
        print("\nPlease install missing dependencies:")
        print("pip install -r requirements.txt")
        sys.exit(1)

    # Test local modules
    if not test_local_modules():
        sys.exit(1)

    # Test basic functionality
    if not test_basic_functionality():
        sys.exit(1)

    print("\n" + "=" * 50)
    print("All tests passed! You can now run:")
    print("python src/main.py")