import sys
import os
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                              QGridLayout, QSplitter, QTabWidget, QGroupBox,
                              QLabel, QLineEdit, QPushButton, QComboBox, QTextEdit,
                              QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem,
                              QProgressBar, QStatusBar, QMenuBar, QToolBar, QFileDialog,
                              QMessageBox, QCheckBox, QSpinBox)
from PySide6.QtCore import Qt, QTimer, QThread, pyqtSignal
from PySide6.QtGui import QAction, QIcon, QFont
from qfluentwidgets import (FluentIcon, setTheme, Theme, FluentWindow, NavigationAvatarWidget,
                           qrouter, SubtitleLabel, setFont, BodyLabel, PushButton,
                           PrimaryPushButton, ComboBox, LineEdit, TextEdit, CheckBox)

from ..core.database_manager import DatabaseManager
from ..core.config_manager import ConfigManager
from ..core.folder_monitor import FolderMonitor
from ..core.email_sender import EmailSenderFactory
from ..core.template_engine import EmailTemplateEngine
from ..utils.logger import setup_logger

class EmailAutomationWorker(QThread):
    """Worker thread for email automation"""
    file_processed = pyqtSignal(str, str, bool)  # file_path, key, success
    error_occurred = pyqtSignal(str)

    def __init__(self, config_manager, database_manager, template_engine):
        super().__init__()
        self.config_manager = config_manager
        self.database_manager = database_manager
        self.template_engine = template_engine
        self.folder_monitor = FolderMonitor()
        self.logger = setup_logger(__name__)

    def process_file(self, file_path: str, key: str):
        """Process detected file"""
        try:
            # Get supplier data
            supplier = self.database_manager.get_supplier_by_key(key)
            if not supplier:
                self.error_occurred.emit(f"Supplier not found for key: {key}")
                return

            # Get current profile config
            profile_config = self.config_manager.get_profile_config()

            # Prepare email content
            variables = self.template_engine.prepare_variables(file_path, supplier)

            # Render subject
            subject = self.template_engine.process_simple_variables(
                profile_config['subject_template'], variables
            )

            # Render body from template file
            body_template_path = profile_config.get('body_template', 'default_template.html')
            try:
                body = self.template_engine.render_file_template(body_template_path, variables)
            except:
                # Fallback to simple template
                body = f"Document {variables['filename']} untuk {supplier['supplier_name']}"

            # Create email sender
            email_sender = EmailSenderFactory.create_sender(
                profile_config['email_client'],
                **{k: v for k, v in profile_config.items() if k.startswith('smtp_')}
            )

            # Send email
            success = email_sender.send_email(
                to_emails=supplier['emails'],
                cc_emails=supplier.get('cc_emails', []),
                bcc_emails=supplier.get('bcc_emails', []),
                subject=subject,
                body=body,
                attachment_path=file_path
            )

            if success:
                # Log to database
                log_data = {
                    'file_path': file_path,
                    'filename': os.path.basename(file_path),
                    'supplier_key': key,
                    'recipient_emails': supplier['emails'],
                    'cc_emails': supplier.get('cc_emails', []),
                    'bcc_emails': supplier.get('bcc_emails', []),
                    'subject': subject,
                    'body': body,
                    'template_used': body_template_path,
                    'email_client': profile_config['email_client'],
                    'status': 'sent'
                }
                self.database_manager.log_email_sent(log_data)

                # Move file to sent folder
                sent_folder = profile_config['sent_folder']
                self.folder_monitor.move_file_to_sent(file_path, sent_folder)

            self.file_processed.emit(file_path, key, success)

        except Exception as e:
            self.logger.error(f"Error processing file {file_path}: {str(e)}")
            self.error_occurred.emit(f"Error processing {os.path.basename(file_path)}: {str(e)}")

class MainWindow(FluentWindow):
    """Main application window"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Automation Desktop")
        self.resize(1200, 800)

        # Initialize core components
        self.config_manager = ConfigManager()
        self.database_manager = DatabaseManager(self.config_manager.get_database_path())
        self.template_engine = EmailTemplateEngine(self.config_manager.get_template_dir())
        self.worker = EmailAutomationWorker(self.config_manager, self.database_manager, self.template_engine)

        # Setup logger
        self.logger = setup_logger(__name__)

        # Initialize UI
        self.init_ui()
        self.init_connections()

        # Status
        self.is_monitoring = False

    def init_ui(self):
        """Initialize user interface"""
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        main_layout = QHBoxLayout(central_widget)

        # Create splitter
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # Left panel - Configuration
        config_panel = self.create_config_panel()
        splitter.addWidget(config_panel)

        # Center panel - Template & Preview
        template_panel = self.create_template_panel()
        splitter.addWidget(template_panel)

        # Right panel - Status & Logs
        status_panel = self.create_status_panel()
        splitter.addWidget(status_panel)

        # Set splitter proportions
        splitter.setSizes([300, 500, 400])

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

        # Load current configuration
        self.load_current_config()

    def create_config_panel(self):
        """Create configuration panel"""
        group = QGroupBox("Configuration")
        layout = QVBoxLayout(group)

        # Profile selection
        layout.addWidget(QLabel("Profile:"))
        self.profile_combo = ComboBox()
        layout.addWidget(self.profile_combo)

        # Monitor folder
        layout.addWidget(QLabel("Monitor Folder:"))
        folder_layout = QHBoxLayout()
        self.monitor_folder_edit = LineEdit()
        self.browse_monitor_btn = PushButton("Browse")
        folder_layout.addWidget(self.monitor_folder_edit)
        folder_layout.addWidget(self.browse_monitor_btn)
        layout.addLayout(folder_layout)

        # Sent folder
        layout.addWidget(QLabel("Sent Folder:"))
        sent_layout = QHBoxLayout()
        self.sent_folder_edit = LineEdit()
        self.browse_sent_btn = PushButton("Browse")
        sent_layout.addWidget(self.sent_folder_edit)
        sent_layout.addWidget(self.browse_sent_btn)
        layout.addLayout(sent_layout)

        # Key pattern
        layout.addWidget(QLabel("Key Pattern (Regex):"))
        self.key_pattern_edit = LineEdit()
        layout.addWidget(self.key_pattern_edit)

        # Email client
        layout.addWidget(QLabel("Email Client:"))
        self.email_client_combo = ComboBox()
        self.email_client_combo.addItems(["outlook", "smtp"])
        layout.addWidget(self.email_client_combo)

        # Control buttons
        button_layout = QVBoxLayout()
        self.start_btn = PrimaryPushButton("Start Monitoring")
        self.stop_btn = PushButton("Stop Monitoring")
        self.stop_btn.setEnabled(False)
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        layout.addLayout(button_layout)

        layout.addStretch()
        return group

    def create_template_panel(self):
        """Create template panel"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Tab widget for templates
        tab_widget = QTabWidget()
        layout.addWidget(tab_widget)

        # Subject template tab
        subject_widget = QWidget()
        subject_layout = QVBoxLayout(subject_widget)
        subject_layout.addWidget(QLabel("Subject Template:"))
        self.subject_template_edit = LineEdit()
        subject_layout.addWidget(self.subject_template_edit)

        # Body template tab
        body_widget = QWidget()
        body_layout = QVBoxLayout(body_widget)
        body_layout.addWidget(QLabel("Body Template:"))
        self.body_template_edit = TextEdit()
        body_layout.addWidget(self.body_template_edit)

        # Preview tab
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        preview_layout.addWidget(QLabel("Preview:"))
        self.preview_text = TextEdit()
        self.preview_text.setReadOnly(True)
        preview_layout.addWidget(self.preview_text)
        self.preview_btn = PushButton("Generate Preview")
        preview_layout.addWidget(self.preview_btn)

        tab_widget.addTab(subject_widget, "Subject")
        tab_widget.addTab(body_widget, "Body")
        tab_widget.addTab(preview_widget, "Preview")

        return widget

    def create_status_panel(self):
        """Create status panel"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Monitoring status
        status_group = QGroupBox("Status")
        status_layout = QVBoxLayout(status_group)

        self.status_label = QLabel("Monitoring: Stopped")
        self.status_label.setStyleSheet("color: red; font-weight: bold;")
        status_layout.addWidget(self.status_label)

        self.files_processed_label = QLabel("Files Processed: 0")
        status_layout.addWidget(self.files_processed_label)

        layout.addWidget(status_group)

        # Recent files
        recent_group = QGroupBox("Recent Files")
        recent_layout = QVBoxLayout(recent_group)

        self.recent_files_list = QListWidget()
        recent_layout.addWidget(self.recent_files_list)

        layout.addWidget(recent_group)

        # Email logs
        logs_group = QGroupBox("Email Logs")
        logs_layout = QVBoxLayout(logs_group)

        self.logs_table = QTableWidget()
        self.logs_table.setColumnCount(4)
        self.logs_table.setHorizontalHeaderLabels(["Time", "File", "Supplier", "Status"])
        logs_layout.addWidget(self.logs_table)

        layout.addWidget(logs_group)

        return widget

    def init_connections(self):
        """Initialize signal connections"""
        # Buttons
        self.start_btn.clicked.connect(self.start_monitoring)
        self.stop_btn.clicked.connect(self.stop_monitoring)
        self.browse_monitor_btn.clicked.connect(self.browse_monitor_folder)
        self.browse_sent_btn.clicked.connect(self.browse_sent_folder)
        self.preview_btn.clicked.connect(self.generate_preview)

        # Profile combo
        self.profile_combo.currentTextChanged.connect(self.load_profile_config)

        # Worker signals
        self.worker.file_processed.connect(self.on_file_processed)
        self.worker.error_occurred.connect(self.on_error_occurred)

    def load_current_config(self):
        """Load current configuration"""
        # Load profiles
        profiles = self.config_manager.get_available_profiles()
        self.profile_combo.clear()
        for profile in profiles:
            self.profile_combo.addItem(profile['display_name'], profile['name'])

        # Set current profile
        current_profile = self.config_manager.get_current_profile()
        for i in range(self.profile_combo.count()):
            if self.profile_combo.itemData(i) == current_profile:
                self.profile_combo.setCurrentIndex(i)
                break

        self.load_profile_config()

    def load_profile_config(self):
        """Load configuration for selected profile"""
        profile_name = self.profile_combo.currentData()
        if not profile_name:
            return

        try:
            config = self.config_manager.get_profile_config(profile_name)

            self.monitor_folder_edit.setText(config.get('monitor_folder', ''))
            self.sent_folder_edit.setText(config.get('sent_folder', ''))
            self.key_pattern_edit.setText(config.get('key_pattern', ''))

            # Set email client
            client = config.get('email_client', 'outlook')
            index = self.email_client_combo.findText(client)
            if index >= 0:
                self.email_client_combo.setCurrentIndex(index)

            self.subject_template_edit.setText(config.get('subject_template', ''))

            # Load body template
            template_file = config.get('body_template', '')
            if template_file:
                try:
                    template_path = os.path.join(self.config_manager.get_template_dir(), template_file)
                    if os.path.exists(template_path):
                        with open(template_path, 'r', encoding='utf-8') as f:
                            self.body_template_edit.setText(f.read())
                except:
                    pass

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load profile: {str(e)}")

    def browse_monitor_folder(self):
        """Browse for monitor folder"""
        folder = QFileDialog.getExistingDirectory(self, "Select Monitor Folder")
        if folder:
            self.monitor_folder_edit.setText(folder)

    def browse_sent_folder(self):
        """Browse for sent folder"""
        folder = QFileDialog.getExistingDirectory(self, "Select Sent Folder")
        if folder:
            self.sent_folder_edit.setText(folder)

    def generate_preview(self):
        """Generate template preview"""
        try:
            subject_template = self.subject_template_edit.text()
            body_template = self.body_template_edit.toPlainText()

            # Use sample data for preview
            preview_html = self.template_engine.preview_template(body_template)
            self.preview_text.setHtml(preview_html)

        except Exception as e:
            QMessageBox.warning(self, "Preview Error", f"Failed to generate preview: {str(e)}")

    def start_monitoring(self):
        """Start folder monitoring"""
        try:
            monitor_folder = self.monitor_folder_edit.text().strip()
            if not monitor_folder or not os.path.exists(monitor_folder):
                QMessageBox.warning(self, "Error", "Please select a valid monitor folder")
                return

            # Save current configuration
            self.save_current_config()

            # Start monitoring
            profile_config = self.config_manager.get_profile_config()
            success = self.worker.folder_monitor.start_monitoring(
                folder_path=monitor_folder,
                callback=self.worker.process_file,
                key_pattern=profile_config['key_pattern'],
                file_extensions=profile_config.get('file_extensions', [])
            )

            if success:
                self.is_monitoring = True
                self.start_btn.setEnabled(False)
                self.stop_btn.setEnabled(True)
                self.status_label.setText("Monitoring: Active")
                self.status_label.setStyleSheet("color: green; font-weight: bold;")
                self.status_bar.showMessage(f"Monitoring folder: {monitor_folder}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to start monitoring: {str(e)}")

    def stop_monitoring(self):
        """Stop folder monitoring"""
        try:
            self.worker.folder_monitor.stop_monitoring()
            self.is_monitoring = False
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.status_label.setText("Monitoring: Stopped")
            self.status_label.setStyleSheet("color: red; font-weight: bold;")
            self.status_bar.showMessage("Monitoring stopped")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to stop monitoring: {str(e)}")

    def save_current_config(self):
        """Save current configuration"""
        try:
            profile_name = self.profile_combo.currentData()
            if not profile_name:
                return

            config_data = {
                'monitor_folder': self.monitor_folder_edit.text(),
                'sent_folder': self.sent_folder_edit.text(),
                'key_pattern': self.key_pattern_edit.text(),
                'email_client': self.email_client_combo.currentText(),
                'subject_template': self.subject_template_edit.text(),
                'body_template': 'custom_template.html'  # Could be made configurable
            }

            self.config_manager.save_profile_config(profile_name, config_data)

        except Exception as e:
            self.logger.error(f"Failed to save configuration: {str(e)}")

    def on_file_processed(self, file_path: str, key: str, success: bool):
        """Handle file processed signal"""
        filename = os.path.basename(file_path)
        status = "Success" if success else "Failed"

        # Add to recent files
        item = QListWidgetItem(f"{filename} ({key}) - {status}")
        if success:
            item.setIcon(FluentIcon.ACCEPT)
        else:
            item.setIcon(FluentIcon.CANCEL)
        self.recent_files_list.insertItem(0, item)

        # Update files processed counter
        files_count = getattr(self, 'files_processed_count', 0) + 1
        self.files_processed_count = files_count
        self.files_processed_label.setText(f"Files Processed: {files_count}")

        # Refresh logs table
        self.refresh_logs_table()

    def on_error_occurred(self, error_message: str):
        """Handle error signal"""
        self.status_bar.showMessage(f"Error: {error_message}", 5000)
        QMessageBox.warning(self, "Processing Error", error_message)

    def refresh_logs_table(self):
        """Refresh email logs table"""
        try:
            logs = self.database_manager.get_email_logs(limit=50)
            self.logs_table.setRowCount(len(logs))

            for row, log in enumerate(logs):
                self.logs_table.setItem(row, 0, QTableWidgetItem(log.get('sent_at', '')))
                self.logs_table.setItem(row, 1, QTableWidgetItem(log.get('filename', '')))
                self.logs_table.setItem(row, 2, QTableWidgetItem(log.get('supplier_key', '')))
                self.logs_table.setItem(row, 3, QTableWidgetItem(log.get('status', '')))

        except Exception as e:
            self.logger.error(f"Failed to refresh logs table: {str(e)}")

    def closeEvent(self, event):
        """Handle window close event"""
        if self.is_monitoring:
            self.stop_monitoring()
        event.accept()