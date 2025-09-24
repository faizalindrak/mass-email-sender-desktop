import sys
import os
import logging

# Add src directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

try:
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtCore import Qt
    print("PySide6 imported successfully")
except ImportError as e:
    print(f"Failed to import PySide6: {e}")
    print("Please install PySide6: pip install PySide6")
    sys.exit(1)

try:
    from qfluentwidgets import setTheme, Theme, setThemeColor
    print("qfluentwidgets imported successfully")
except ImportError as e:
    print(f"Failed to import qfluentwidgets: {e}")
    print("Please install PyQt-Fluent-Widgets: pip install PyQt-Fluent-Widgets")
    sys.exit(1)

try:
    from ui.main_window import MainWindow
    from core.config_manager import ConfigManager
    from core.template_engine import EmailTemplateEngine
    from utils.logger import setup_logger
    print("Local modules imported successfully")
except ImportError as e:
    print(f"Failed to import local modules: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

def main():
    """Main application entry point"""
    print("Starting main function...")

    try:
        # Create application first
        print("Creating QApplication...")
        app = QApplication(sys.argv)
        app.setApplicationName("Email Automation Desktop")
        app.setApplicationVersion("1.0.0")
        app.setOrganizationName("Email Automation")
        
        # Set application icon
        try:
            from PySide6.QtGui import QIcon
            icon_path = os.path.join(os.path.dirname(os.path.dirname(current_dir)), 'icon.ico')
            if os.path.exists(icon_path):
                app.setWindowIcon(QIcon(icon_path))
                print(f"Application icon set: {icon_path}")
            else:
                print(f"Icon file not found: {icon_path}")
        except Exception as e:
            print(f"Failed to set application icon: {e}")

        # Enable high DPI scaling (remove deprecated warnings)
        # These are deprecated in newer Qt versions but still needed for older versions
        try:
            QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
        except:
            pass

        # Setup theme
        print("Setting up theme...")
        setTheme(Theme.AUTO)
        setThemeColor('#0078d4')

        print("Initializing configuration...")
        # Initialize configuration
        config_manager = ConfigManager()

        # Setup logging
        print("Setting up logging...")
        log_config = config_manager.get_log_config()
        logger = setup_logger(
            'email_automation',
            log_file=log_config['file'],
            level=getattr(logging, log_config['level'].upper(), logging.INFO)
        )

        logger.info("Starting Email Automation Desktop application")

        # Create default templates if they don't exist
        print("Creating default templates...")
        template_engine = EmailTemplateEngine(config_manager.get_template_dir())
        template_engine.create_default_templates()

        # Create and show main window (after QApplication is created)
        print("Creating main window...")
        window = MainWindow()
        print("Showing main window...")
        window.show()

        # Auto-start monitoring if configured
        if config_manager.should_auto_start_monitoring():
            logger.info("Auto-starting monitoring")

        logger.info("Application started successfully")
        print("Application started successfully")

        # Run application
        print("Starting application event loop...")
        exit_code = app.exec()
        logger.info(f"Application exited with code: {exit_code}")
        return exit_code

    except Exception as e:
        print(f"Failed to start application: {str(e)}")
        import traceback
        traceback.print_exc()

        # Try to show error dialog if QApplication exists
        try:
            if 'app' in locals():
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.critical(None, "Error", f"Failed to start application:\n{str(e)}")
        except:
            pass

        return 1

if __name__ == "__main__":
    import logging
    sys.exit(main())