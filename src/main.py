import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from qfluentwidgets import setTheme, Theme, setThemeColor

# Add src directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from ui.main_window import MainWindow
from core.config_manager import ConfigManager
from core.template_engine import EmailTemplateEngine
from utils.logger import setup_logger

def main():
    """Main application entry point"""

    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    # Create application
    app = QApplication(sys.argv)
    app.setApplicationName("Email Automation Desktop")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("Email Automation")

    # Setup theme
    setTheme(Theme.AUTO)
    setThemeColor('#0078d4')

    try:
        # Initialize configuration
        config_manager = ConfigManager()

        # Setup logging
        log_config = config_manager.get_log_config()
        logger = setup_logger(
            'email_automation',
            log_file=log_config['file'],
            level=getattr(logging, log_config['level'].upper(), logging.INFO)
        )

        logger.info("Starting Email Automation Desktop application")

        # Create default templates if they don't exist
        template_engine = EmailTemplateEngine(config_manager.get_template_dir())
        template_engine.create_default_templates()

        # Create and show main window
        window = MainWindow()
        window.show()

        # Auto-start monitoring if configured
        if config_manager.should_auto_start_monitoring():
            logger.info("Auto-starting monitoring")
            # Note: This could be implemented in MainWindow's showEvent

        logger.info("Application started successfully")

        # Run application
        exit_code = app.exec()
        logger.info(f"Application exited with code: {exit_code}")
        return exit_code

    except Exception as e:
        print(f"Failed to start application: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    import logging
    sys.exit(main())