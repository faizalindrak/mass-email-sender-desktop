import sys
import os
import time
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                              QGridLayout, QSplitter, QTabWidget,
                              QTableWidget, QTableWidgetItem, QListWidget, QListWidgetItem,
                              QFileDialog, QMessageBox)
from PySide6.QtCore import Qt, QTimer, QThread, Signal
from PySide6.QtGui import QAction, QIcon, QFont

from qfluentwidgets import (FluentIcon, setTheme, Theme, FluentWindow, NavigationItemPosition,
                           qrouter, SubtitleLabel, setFont, BodyLabel, PushButton, TitleLabel,
                           PrimaryPushButton, ComboBox, LineEdit, TextEdit, CheckBox,
                           CardWidget, SimpleCardWidget, HeaderCardWidget, GroupHeaderCardWidget,
                           SwitchButton, ToggleButton, Pivot, PivotItem, ScrollArea,
                           InfoBar, InfoBarPosition, StrongBodyLabel, CaptionLabel,
                           TableWidget, ListWidget, TreeWidget, ProgressBar,
                           ToolTip, TeachingTip, TeachingTipTailPosition, PopupTeachingTip,
                           FlyoutViewBase, Flyout, FlyoutAnimationType)

from core.database_manager import DatabaseManager  # Changed relative import
from core.config_manager import ConfigManager
from core.folder_monitor import FolderMonitor
from core.email_sender import EmailSenderFactory
from core.template_engine import EmailTemplateEngine
from utils.logger import setup_logger

class EmailAutomationWorker(QThread):
    """Worker thread for email automation"""
    file_processed = Signal(str, str, bool)  # file_path, key, success
    error_occurred = Signal(str)

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

            # Build custom variables from profile (default CC/BCC + custom vars)
            custom_vars = {}
            # Default CC/BCC as variables (join for display in templates)
            default_cc_list = [e.strip() for e in profile_config.get('default_cc', '').split(';') if e.strip()]
            default_bcc_list = [e.strip() for e in profile_config.get('default_bcc', '').split(';') if e.strip()]
            if default_cc_list:
                custom_vars['default_cc'] = default_cc_list
            if default_bcc_list:
                custom_vars['default_bcc'] = default_bcc_list

            # Custom named variables
            c1n = profile_config.get('custom1_name', '').strip()
            c1v = profile_config.get('custom1_value', '')
            if c1n:
                custom_vars[c1n] = c1v
            c2n = profile_config.get('custom2_name', '').strip()
            c2v = profile_config.get('custom2_value', '')
            if c2n:
                custom_vars[c2n] = c2v

            # Prepare email content
            variables = self.template_engine.prepare_variables(file_path, supplier, custom_vars)

            # Resolve subject template: prefer email_form.subject if present; fallback to profile subject_template
            email_form = profile_config.get('email_form', {}) or {}
            subject_template = email_form.get('subject') or profile_config.get('subject_template', '[filename_without_ext]')
            subject = self.template_engine.process_simple_variables(
                subject_template,
                variables
            )

            # Render body from template file
            body_template_path = profile_config.get('body_template', 'default_template.html')
            try:
                body = self.template_engine.render_file_template(body_template_path, variables)
            except Exception:
                # Fallback to simple template
                body = f"Document {variables['filename']} untuk {supplier['supplier_name']}"

            # Create email sender
            email_sender = EmailSenderFactory.create_sender(
                profile_config['email_client'],
                **{k: v for k, v in profile_config.items() if k.startswith('smtp_')}
            )

            # Merge supplier CC/BCC with default CC/BCC from profile
            merged_cc = (supplier.get('cc_emails', []) or []) + default_cc_list
            merged_bcc = (supplier.get('bcc_emails', []) or []) + default_bcc_list
            
            # Debug logging for CC/BCC
            self.logger.info(f"Default CC from profile: {default_cc_list}")
            self.logger.info(f"Default BCC from profile: {default_bcc_list}")
            self.logger.info(f"Supplier CC: {supplier.get('cc_emails', [])}")
            self.logger.info(f"Supplier BCC: {supplier.get('bcc_emails', [])}")
            self.logger.info(f"Merged CC: {merged_cc}")
            self.logger.info(f"Merged BCC: {merged_bcc}")

            # Send email
            success = email_sender.send_email(
                to_emails=supplier['emails'],
                cc_emails=merged_cc,
                bcc_emails=merged_bcc,
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
                    'cc_emails': merged_cc,
                    'bcc_emails': merged_bcc,
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
        finally:
            # Ensure a 2-second delay between processing/sending files
            try:
                time.sleep(2)
            except Exception:
                pass

class MainWindow(FluentWindow):
    """Main application window with Fluent Design"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Automation Desktop")
        self.resize(1200, 800)
        
        # Set fluent theme
        setTheme(Theme.AUTO)

        # Initialize core components
        self.config_manager = ConfigManager()
        self.database_manager = DatabaseManager(self.config_manager.get_database_path())
        self.template_engine = EmailTemplateEngine(self.config_manager.get_template_dir())
        self.worker = EmailAutomationWorker(self.config_manager, self.database_manager, self.template_engine)

        # Setup logger
        self.logger = setup_logger(__name__)

        # Status
        self.is_monitoring = False
        self.files_processed_count = 0

        # Initialize UI
        self.init_ui()
        self.init_connections()

    def init_ui(self):
        """Initialize user interface with Fluent Design"""
        # Create navigation interfaces
        self.config_interface = self.create_config_interface()
        self.config_interface.setObjectName('ConfigInterface')
        
        self.template_interface = self.create_template_interface()
        self.template_interface.setObjectName('TemplateInterface')
        
        self.status_interface = self.create_status_interface()
        self.status_interface.setObjectName('StatusInterface')
        
        # Add navigation items
        self.addSubInterface(
            self.config_interface,
            FluentIcon.SETTING,
            'Configuration',
            NavigationItemPosition.TOP
        )
        
        self.addSubInterface(
            self.template_interface,
            FluentIcon.EDIT,
            'Templates',
            NavigationItemPosition.TOP
        )
        
        self.addSubInterface(
            self.status_interface,
            FluentIcon.INFO,
            'Status & Logs',
            NavigationItemPosition.TOP
        )

        # Load current configuration
        self.load_current_config()
        
        # Show configuration interface by default
        self.stackedWidget.setCurrentWidget(self.config_interface)

    def create_config_interface(self):
        """Create configuration interface with Fluent Design"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Title
        title_label = TitleLabel("Email Automation Configuration")
        layout.addWidget(title_label)

        # Profile Management Card
        profile_card = GroupHeaderCardWidget("Profile Management")
        profile_layout = QVBoxLayout()

        # Profile selection row
        profile_row = QHBoxLayout()
        profile_row.addWidget(BodyLabel("Current Profile:"))
        self.profile_combo = ComboBox()
        self.load_profile_btn = PushButton(FluentIcon.FOLDER, "Load")
        self.save_profile_btn = PrimaryPushButton(FluentIcon.SAVE, "Save")
        profile_row.addWidget(self.profile_combo)
        profile_row.addWidget(self.load_profile_btn)
        profile_row.addWidget(self.save_profile_btn)
        profile_row.addStretch()
        profile_layout.addLayout(profile_row)
        
        profile_card.viewLayout.addLayout(profile_layout)
        layout.addWidget(profile_card)

        # System Paths Card
        paths_card = GroupHeaderCardWidget("System Paths")
        paths_layout = QVBoxLayout()

        # Database file
        db_row = QHBoxLayout()
        db_row.addWidget(StrongBodyLabel("Database File:"))
        db_row.addStretch()
        paths_layout.addLayout(db_row)
        
        db_input_row = QHBoxLayout()
        self.database_path_edit = LineEdit()
        self.database_path_edit.setPlaceholderText("Select database file...")
        self.browse_database_btn = PushButton(FluentIcon.FOLDER, "Browse")
        db_input_row.addWidget(self.database_path_edit)
        db_input_row.addWidget(self.browse_database_btn)
        paths_layout.addLayout(db_input_row)

        # Template folder
        tpl_row = QHBoxLayout()
        tpl_row.addWidget(StrongBodyLabel("Template Folder:"))
        tpl_row.addStretch()
        paths_layout.addLayout(tpl_row)
        
        tpl_input_row = QHBoxLayout()
        self.template_dir_edit = LineEdit()
        self.template_dir_edit.setPlaceholderText("Select template folder...")
        self.browse_template_btn = PushButton(FluentIcon.FOLDER, "Browse")
        tpl_input_row.addWidget(self.template_dir_edit)
        tpl_input_row.addWidget(self.browse_template_btn)
        paths_layout.addLayout(tpl_input_row)
        
        paths_card.viewLayout.addLayout(paths_layout)
        layout.addWidget(paths_card)

        # Monitoring Settings Card
        monitoring_card = GroupHeaderCardWidget("Monitoring Settings")
        monitoring_layout = QVBoxLayout()

        # Monitor folder
        monitor_row = QHBoxLayout()
        monitor_row.addWidget(StrongBodyLabel("Monitor Folder:"))
        monitor_row.addStretch()
        monitoring_layout.addLayout(monitor_row)
        
        monitor_input_row = QHBoxLayout()
        self.monitor_folder_edit = LineEdit()
        self.monitor_folder_edit.setPlaceholderText("Select folder to monitor...")
        self.browse_monitor_btn = PushButton(FluentIcon.FOLDER, "Browse")
        monitor_input_row.addWidget(self.monitor_folder_edit)
        monitor_input_row.addWidget(self.browse_monitor_btn)
        monitoring_layout.addLayout(monitor_input_row)

        # Sent folder
        sent_row = QHBoxLayout()
        sent_row.addWidget(StrongBodyLabel("Sent Folder:"))
        sent_row.addStretch()
        monitoring_layout.addLayout(sent_row)
        
        sent_input_row = QHBoxLayout()
        self.sent_folder_edit = LineEdit()
        self.sent_folder_edit.setPlaceholderText("Select sent files folder...")
        self.browse_sent_btn = PushButton(FluentIcon.FOLDER, "Browse")
        sent_input_row.addWidget(self.sent_folder_edit)
        sent_input_row.addWidget(self.browse_sent_btn)
        monitoring_layout.addLayout(sent_input_row)

        # Key pattern
        pattern_row = QHBoxLayout()
        pattern_row.addWidget(StrongBodyLabel("Key Pattern (Regex):"))
        pattern_row.addStretch()
        monitoring_layout.addLayout(pattern_row)
        
        self.key_pattern_edit = LineEdit()
        self.key_pattern_edit.setPlaceholderText("Enter regex pattern to extract keys from filenames...")
        monitoring_layout.addWidget(self.key_pattern_edit)

        # Email client
        client_row = QHBoxLayout()
        client_row.addWidget(StrongBodyLabel("Email Client:"))
        client_row.addStretch()
        monitoring_layout.addLayout(client_row)
        
        self.email_client_combo = ComboBox()
        self.email_client_combo.addItems(["outlook", "thunderbird", "smtp"])
        monitoring_layout.addWidget(self.email_client_combo)

        monitoring_card.viewLayout.addLayout(monitoring_layout)
        layout.addWidget(monitoring_card)

        # File Extensions Card
        extensions_card = GroupHeaderCardWidget("File Types to Monitor")
        extensions_layout = QVBoxLayout()

        # Extension controls
        ext_controls_layout = QHBoxLayout()
        self.scan_extensions_btn = PushButton(FluentIcon.SEARCH, "Scan Extensions")
        self.select_all_ext_btn = PushButton(FluentIcon.CHECKBOX, "Select All")
        self.clear_ext_btn = PushButton(FluentIcon.CANCEL, "Clear")
        ext_controls_layout.addWidget(self.scan_extensions_btn)
        ext_controls_layout.addWidget(self.select_all_ext_btn)
        ext_controls_layout.addWidget(self.clear_ext_btn)
        ext_controls_layout.addStretch()
        extensions_layout.addLayout(ext_controls_layout)

        self.extensions_list = ListWidget()
        extensions_layout.addWidget(self.extensions_list)
        
        extensions_card.viewLayout.addLayout(extensions_layout)
        layout.addWidget(extensions_card)

        # Variables Card
        variables_card = GroupHeaderCardWidget("Constant Variables")
        variables_layout = QVBoxLayout()

        # Default CC emails
        cc_row = QHBoxLayout()
        cc_row.addWidget(StrongBodyLabel("Default CC:"))
        cc_row.addStretch()
        variables_layout.addLayout(cc_row)
        
        self.default_cc_edit = LineEdit()
        self.default_cc_edit.setPlaceholderText("Default CC emails (semicolon separated)")
        variables_layout.addWidget(self.default_cc_edit)

        # Default BCC emails
        bcc_row = QHBoxLayout()
        bcc_row.addWidget(StrongBodyLabel("Default BCC:"))
        bcc_row.addStretch()
        variables_layout.addLayout(bcc_row)
        
        self.default_bcc_edit = LineEdit()
        self.default_bcc_edit.setPlaceholderText("Default BCC emails (semicolon separated)")
        variables_layout.addWidget(self.default_bcc_edit)

        # Custom variable 1
        custom1_row = QHBoxLayout()
        custom1_row.addWidget(StrongBodyLabel("Custom Variable 1:"))
        custom1_row.addStretch()
        variables_layout.addLayout(custom1_row)
        
        custom1_layout = QHBoxLayout()
        self.custom1_name_edit = LineEdit()
        self.custom1_name_edit.setPlaceholderText("Variable name")
        self.custom1_value_edit = LineEdit()
        self.custom1_value_edit.setPlaceholderText("Variable value")
        custom1_layout.addWidget(self.custom1_name_edit)
        custom1_layout.addWidget(self.custom1_value_edit)
        variables_layout.addLayout(custom1_layout)

        # Custom variable 2
        custom2_row = QHBoxLayout()
        custom2_row.addWidget(StrongBodyLabel("Custom Variable 2:"))
        custom2_row.addStretch()
        variables_layout.addLayout(custom2_row)
        
        custom2_layout = QHBoxLayout()
        self.custom2_name_edit = LineEdit()
        self.custom2_name_edit.setPlaceholderText("Variable name")
        self.custom2_value_edit = LineEdit()
        self.custom2_value_edit.setPlaceholderText("Variable value")
        custom2_layout.addWidget(self.custom2_name_edit)
        custom2_layout.addWidget(self.custom2_value_edit)
        variables_layout.addLayout(custom2_layout)

        variables_card.viewLayout.addLayout(variables_layout)
        layout.addWidget(variables_card)

        # Control Card
        control_card = SimpleCardWidget()
        control_layout = QVBoxLayout(control_card)
        control_layout.setContentsMargins(20, 20, 20, 20)
        
        self.toggle_monitoring_btn = PrimaryPushButton(FluentIcon.PLAY, "Start Monitoring")
        self.toggle_monitoring_btn.setFixedHeight(40)
        control_layout.addWidget(self.toggle_monitoring_btn)

        layout.addWidget(control_card)
        layout.addStretch()

        widget.setWidget(content_widget)
        return widget

    def create_template_interface(self):
        """Create template interface with Fluent Design"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Title
        title_label = TitleLabel("Email Templates & Composition")
        layout.addWidget(title_label)

        # Pivot for different sections
        pivot = Pivot()
        pivot.addItem(
            routeKey='email_form',
            text='Email Form',
            onClick=lambda: self.stackedWidget_template.setCurrentIndex(0)
        )
        pivot.addItem(
            routeKey='variables',
            text='Variables',
            onClick=lambda: self.stackedWidget_template.setCurrentIndex(1)
        )
        pivot.addItem(
            routeKey='preview',
            text='Preview',
            onClick=lambda: self.stackedWidget_template.setCurrentIndex(2)
        )
        layout.addWidget(pivot)

        # Create stacked widget for pivot content
        from PySide6.QtWidgets import QStackedWidget
        self.stackedWidget_template = QStackedWidget()
        
        # Email Form Content
        email_form_widget = self.create_email_form_content()
        self.stackedWidget_template.addWidget(email_form_widget)
        
        # Variables Content
        variables_widget = self.create_variables_content()
        self.stackedWidget_template.addWidget(variables_widget)
        
        # Preview Content
        preview_widget = self.create_preview_content()
        self.stackedWidget_template.addWidget(preview_widget)
        
        layout.addWidget(self.stackedWidget_template)

        widget.setWidget(content_widget)
        return widget

    def create_email_form_content(self):
        """Create email form content with fluent cards"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(16)

        # Recipients Card
        recipients_card = GroupHeaderCardWidget("Email Recipients")
        recipients_layout = QVBoxLayout()

        # To field
        to_row = QHBoxLayout()
        to_row.addWidget(StrongBodyLabel("To:"))
        to_row.addStretch()
        recipients_layout.addLayout(to_row)
        self.to_emails_edit = LineEdit()
        self.to_emails_edit.setPlaceholderText("Enter email addresses separated by semicolons")
        recipients_layout.addWidget(self.to_emails_edit)

        # CC field
        cc_row = QHBoxLayout()
        cc_row.addWidget(StrongBodyLabel("CC:"))
        cc_row.addStretch()
        recipients_layout.addLayout(cc_row)
        self.cc_emails_edit = LineEdit()
        self.cc_emails_edit.setPlaceholderText("Enter CC email addresses separated by semicolons")
        recipients_layout.addWidget(self.cc_emails_edit)

        # BCC field
        bcc_row = QHBoxLayout()
        bcc_row.addWidget(StrongBodyLabel("BCC:"))
        bcc_row.addStretch()
        recipients_layout.addLayout(bcc_row)
        self.bcc_emails_edit = LineEdit()
        self.bcc_emails_edit.setPlaceholderText("Enter BCC email addresses separated by semicolons")
        recipients_layout.addWidget(self.bcc_emails_edit)

        recipients_card.viewLayout.addLayout(recipients_layout)
        layout.addWidget(recipients_card)

        # Content Card
        content_card = GroupHeaderCardWidget("Email Content")
        content_layout = QVBoxLayout()

        # Subject field
        subject_row = QHBoxLayout()
        subject_row.addWidget(StrongBodyLabel("Subject:"))
        subject_row.addStretch()
        content_layout.addLayout(subject_row)
        self.email_subject_edit = LineEdit()
        self.email_subject_edit.setPlaceholderText("Enter email subject...")
        content_layout.addWidget(self.email_subject_edit)

        # Template selection
        template_row = QHBoxLayout()
        template_row.addWidget(StrongBodyLabel("Template:"))
        template_row.addStretch()
        content_layout.addLayout(template_row)
        self.template_combo = ComboBox()
        self.load_templates()
        content_layout.addWidget(self.template_combo)

        # Body field
        body_row = QHBoxLayout()
        body_row.addWidget(StrongBodyLabel("Body:"))
        body_row.addStretch()
        content_layout.addLayout(body_row)
        self.email_body_edit = TextEdit()
        self.email_body_edit.setMinimumHeight(200)
        content_layout.addWidget(self.email_body_edit)

        content_card.viewLayout.addLayout(content_layout)
        layout.addWidget(content_card)

        # Actions Card
        actions_card = SimpleCardWidget()
        actions_layout = QHBoxLayout(actions_card)
        actions_layout.setContentsMargins(20, 20, 20, 20)
        
        self.save_template_btn = PushButton(FluentIcon.SAVE, "Save Template")
        self.send_test_email_btn = PrimaryPushButton(FluentIcon.SEND, "Send Test Email")
        actions_layout.addWidget(self.save_template_btn)
        actions_layout.addWidget(self.send_test_email_btn)
        actions_layout.addStretch()

        layout.addWidget(actions_card)
        layout.addStretch()

        widget.setWidget(content)
        return widget

    def create_variables_content(self):
        """Create variables content with fluent cards"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(16)

        # Available Variables Card
        variables_card = GroupHeaderCardWidget("Available Variables")
        variables_layout = QVBoxLayout()

        self.variables_list = ListWidget()
        variables_layout.addWidget(self.variables_list)

        # Variable insertion buttons
        var_buttons_layout = QHBoxLayout()
        self.insert_var_btn = PushButton(FluentIcon.ADD, "Insert to Subject")
        self.insert_var_body_btn = PushButton(FluentIcon.ADD, "Insert to Body")
        var_buttons_layout.addWidget(self.insert_var_btn)
        var_buttons_layout.addWidget(self.insert_var_body_btn)
        var_buttons_layout.addStretch()
        variables_layout.addLayout(var_buttons_layout)

        variables_card.viewLayout.addLayout(variables_layout)
        layout.addWidget(variables_card)

        # Sample Data Card
        sample_card = GroupHeaderCardWidget("Sample Data Preview")
        sample_layout = QVBoxLayout()

        self.sample_data_text = TextEdit()
        self.sample_data_text.setReadOnly(True)
        self.sample_data_text.setMaximumHeight(200)
        sample_layout.addWidget(self.sample_data_text)
        
        # Load variables after sample_data_text is initialized
        self.load_available_variables()

        sample_card.viewLayout.addLayout(sample_layout)
        layout.addWidget(sample_card)
        layout.addStretch()

        widget.setWidget(content)
        return widget

    def create_preview_content(self):
        """Create preview content with fluent cards"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(16)

        # Preview Card
        preview_card = GroupHeaderCardWidget("Email Preview")
        preview_layout = QVBoxLayout()

        self.preview_text = TextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(300)
        preview_layout.addWidget(self.preview_text)

        self.preview_btn = PrimaryPushButton(FluentIcon.VIEW, "Generate Preview")
        preview_layout.addWidget(self.preview_btn)

        preview_card.viewLayout.addLayout(preview_layout)
        layout.addWidget(preview_card)
        layout.addStretch()

        widget.setWidget(content)
        return widget

    def create_status_interface(self):
        """Create status interface with Fluent Design"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Title
        title_label = TitleLabel("Status & Monitoring")
        layout.addWidget(title_label)

        # Status Card
        status_card = GroupHeaderCardWidget("Monitoring Status")
        status_layout = QVBoxLayout()

        # Status indicators in a grid
        status_grid = QHBoxLayout()
        
        # Monitoring status indicator
        monitoring_status_card = SimpleCardWidget()
        monitoring_layout = QVBoxLayout(monitoring_status_card)
        monitoring_layout.setContentsMargins(16, 16, 16, 16)
        
        self.status_label = StrongBodyLabel("Monitoring: Stopped")
        self.status_label.setStyleSheet("color: #d73527; font-weight: bold;")
        monitoring_layout.addWidget(self.status_label)
        
        status_grid.addWidget(monitoring_status_card)
        
        # Files processed counter
        files_counter_card = SimpleCardWidget()
        files_layout = QVBoxLayout(files_counter_card)
        files_layout.setContentsMargins(16, 16, 16, 16)
        
        self.files_processed_label = StrongBodyLabel("Files Processed: 0")
        files_layout.addWidget(self.files_processed_label)
        
        status_grid.addWidget(files_counter_card)
        
        status_layout.addLayout(status_grid)
        status_card.viewLayout.addLayout(status_layout)
        layout.addWidget(status_card)

        # Recent Files Card
        recent_card = GroupHeaderCardWidget("Recent Files")
        recent_layout = QVBoxLayout()

        self.recent_files_list = ListWidget()
        self.recent_files_list.setMinimumHeight(200)
        recent_layout.addWidget(self.recent_files_list)

        recent_card.viewLayout.addLayout(recent_layout)
        layout.addWidget(recent_card)

        # Email Logs Card
        logs_card = GroupHeaderCardWidget("Email Logs")
        logs_layout = QVBoxLayout()

        self.logs_table = TableWidget()
        self.logs_table.setColumnCount(4)
        self.logs_table.setHorizontalHeaderLabels(["Time", "File", "Supplier", "Status"])
        self.logs_table.setMinimumHeight(300)
        logs_layout.addWidget(self.logs_table)

        logs_card.viewLayout.addLayout(logs_layout)
        layout.addWidget(logs_card)

        layout.addStretch()
        widget.setWidget(content_widget)
        return widget

    def init_connections(self):
        """Initialize signal connections"""
        # Buttons
        self.toggle_monitoring_btn.clicked.connect(self.toggle_monitoring)
        self.browse_monitor_btn.clicked.connect(self.browse_monitor_folder)
        self.browse_sent_btn.clicked.connect(self.browse_sent_folder)
        self.browse_database_btn.clicked.connect(self.browse_database_file)
        self.browse_template_btn.clicked.connect(self.browse_template_folder)
        self.save_template_btn.clicked.connect(self.save_template_file)
        self.preview_btn.clicked.connect(self.generate_preview)
        self.load_profile_btn.clicked.connect(self.load_profile_from_file)
        self.save_profile_btn.clicked.connect(self.save_profile_to_file)
        self.send_test_email_btn.clicked.connect(self.send_test_email)
        self.template_combo.currentTextChanged.connect(self.load_selected_template)
        self.insert_var_btn.clicked.connect(self.insert_variable_to_subject)
        self.insert_var_body_btn.clicked.connect(self.insert_variable_to_body)
        # Extensions controls
        self.scan_extensions_btn.clicked.connect(self.scan_monitor_folder_extensions)
        self.select_all_ext_btn.clicked.connect(self.select_all_extensions)
        self.clear_ext_btn.clicked.connect(self.clear_extensions_selection)
        
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

            # Database path (from global config if not in profile)
            db_path = config.get('database_path', self.config_manager.get_database_path())
            self.database_path_edit.setText(db_path)
            # Template dir (from global config)
            self.template_dir_edit.setText(self.config_manager.get_template_dir())
            
            # Other folders and pattern
            self.monitor_folder_edit.setText(config.get('monitor_folder', ''))
            self.sent_folder_edit.setText(config.get('sent_folder', ''))
            self.key_pattern_edit.setText(config.get('key_pattern', ''))

            # Set email client
            client = config.get('email_client', 'outlook')
            index = self.email_client_combo.findText(client)
            if index >= 0:
                self.email_client_combo.setCurrentIndex(index)

            # Populate file extensions from monitoring folder and preselect from config
            try:
                self.scan_monitor_folder_extensions(preselected=config.get('file_extensions', []))
            except Exception as _:
                pass

            # Load constant variables
            self.default_cc_edit.setText(config.get('default_cc', ''))
            self.default_bcc_edit.setText(config.get('default_bcc', ''))
            self.custom1_name_edit.setText(config.get('custom1_name', ''))
            self.custom1_value_edit.setText(config.get('custom1_value', ''))
            self.custom2_name_edit.setText(config.get('custom2_name', ''))
            self.custom2_value_edit.setText(config.get('custom2_value', ''))

            # Load email form fields
            email_form = config.get('email_form', {})
            # Join lists into semicolon strings for UI
            self.to_emails_edit.setText('; '.join(email_form.get('to_emails', [])))
            self.cc_emails_edit.setText('; '.join(email_form.get('cc_emails', [])))
            self.bcc_emails_edit.setText('; '.join(email_form.get('bcc_emails', [])))
            # Subject: prefer stored email_form subject, fallback to subject_template
            subject_text = email_form.get('subject', config.get('subject_template', ''))
            self.email_subject_edit.setText(subject_text)

            # Template selection and body content
            selected_template = email_form.get('selected_template', config.get('body_template', ''))
            if selected_template:
                # Ensure templates are loaded into combo
                self.load_templates()
                idx = self.template_combo.findText(selected_template)
                if idx >= 0:
                    self.template_combo.setCurrentIndex(idx)
                else:
                    # Add missing template name if not present
                    self.template_combo.addItem(selected_template)
                    self.template_combo.setCurrentIndex(self.template_combo.count() - 1)

            # Load body from selected template file (do not persist content in configuration)
            try:
                if selected_template:
                    template_dir = self.template_dir_edit.text().strip() or self.config_manager.get_template_dir()
                    template_path = os.path.join(template_dir, selected_template)
                    if os.path.exists(template_path):
                        with open(template_path, 'r', encoding='utf-8') as f:
                            self.email_body_edit.setPlainText(f.read())
                    else:
                        self.email_body_edit.setPlainText("")
            except Exception as _:
                pass

            # Load variables list
            self.load_available_variables()

            # Reinitialize DatabaseManager if path changed
            try:
                abs_db_path = os.path.abspath(db_path)
                if abs_db_path != getattr(self.database_manager, 'db_path', None):
                    self.database_manager = DatabaseManager(abs_db_path)
                    self.worker.database_manager = self.database_manager
                    self.refresh_logs_table()
            except Exception as db_e:
                self.logger.error(f"Failed to initialize database: {str(db_e)}")

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load profile: {str(e)}")

    def browse_monitor_folder(self):
        """Browse for monitor folder"""
        folder = QFileDialog.getExistingDirectory(self, "Select Monitor Folder")
        if folder:
            self.monitor_folder_edit.setText(folder)
            # Auto-set sent folder to be inside monitor folder if not already set
            if not self.sent_folder_edit.text():
                sent_folder = os.path.join(folder, "sent")
                self.sent_folder_edit.setText(sent_folder)
            # Scan and populate available file extensions from selected folder
            try:
                self.scan_monitor_folder_extensions(preselected=self.config_manager.get_profile_config().get('file_extensions', []))
            except Exception as _:
                pass

    def browse_sent_folder(self):
        """Browse for sent folder"""
        monitor_folder = self.monitor_folder_edit.text().strip()

        # Default to sent folder inside monitor folder
        default_sent = os.path.join(monitor_folder, "sent") if monitor_folder else ""

        folder = QFileDialog.getExistingDirectory(self, "Select Sent Folder", default_sent)
        if folder:
            self.sent_folder_edit.setText(folder)
        elif monitor_folder and not self.sent_folder_edit.text():
            # Auto-create sent folder inside monitor folder if not selected
            sent_folder = os.path.join(monitor_folder, "sent")
            self.sent_folder_edit.setText(sent_folder)

    def scan_monitor_folder_extensions(self, preselected=None):
        """Scan the monitoring folder and populate available file extensions"""
        try:
            monitor_folder = self.monitor_folder_edit.text().strip()
            extensions = set()
            if monitor_folder and os.path.exists(monitor_folder):
                for filename in os.listdir(monitor_folder):
                    file_path = os.path.join(monitor_folder, filename)
                    if os.path.isfile(file_path):
                        ext = os.path.splitext(filename)[1].lower()
                        if ext:
                            extensions.add(ext)
            # Fallback defaults if no extensions found
            if not extensions:
                extensions = {'.pdf', '.xlsx', '.docx', '.txt'}
            self.update_extensions_list(sorted(extensions), preselected or [])
        except Exception as e:
            self.logger.error(f"Failed to scan extensions: {str(e)}")

    def update_extensions_list(self, extensions, selected):
        """Update the extensions list with checkable items and preselected values"""
        try:
            self.extensions_list.clear()
            for ext in extensions:
                item = QListWidgetItem(ext)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked if ext in selected else Qt.Unchecked)
                self.extensions_list.addItem(item)
        except Exception as e:
            self.logger.error(f"Failed to update extensions list: {str(e)}")

    def get_selected_extensions(self):
        """Get currently selected extensions from the list"""
        selected = []
        try:
            for i in range(self.extensions_list.count()):
                item = self.extensions_list.item(i)
                if item.checkState() == Qt.Checked:
                    selected.append(item.text())
        except Exception as e:
            self.logger.error(f"Failed to read selected extensions: {str(e)}")
        return selected

    def select_all_extensions(self):
        """Select all extensions in the list"""
        try:
            for i in range(self.extensions_list.count()):
                self.extensions_list.item(i).setCheckState(Qt.Checked)
        except Exception as e:
            self.logger.error(f"Failed to select all extensions: {str(e)}")

    def clear_extensions_selection(self):
        """Clear all selections in the extensions list"""
        try:
            for i in range(self.extensions_list.count()):
                self.extensions_list.item(i).setCheckState(Qt.Unchecked)
        except Exception as e:
            self.logger.error(f"Failed to clear extensions selection: {str(e)}")

    def browse_database_file(self):
        """Browse for database file"""
        initial_dir = os.path.dirname(self.database_path_edit.text().strip() or os.path.abspath("database"))
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Database File",
            initial_dir,
            "SQLite Database (*.db);;All files (*.*)"
        )
        if file_path:
            self.database_path_edit.setText(file_path)
            # Reinitialize database manager with new path and update worker
            try:
                self.database_manager = DatabaseManager(file_path)
                self.worker.database_manager = self.database_manager
                self.refresh_logs_table()
                # Persist database path into global JSON config
                self.config_manager.set_database_path(file_path)
                InfoBar.success(
                    title="Database Updated",
                    content=f"Database set to: {file_path}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to open database: {str(e)}")

    def browse_template_folder(self):
        """Browse for template folder"""
        initial_dir = self.template_dir_edit.text().strip() or self.config_manager.get_template_dir()
        folder = QFileDialog.getExistingDirectory(self, "Select Template Folder", initial_dir)
        if folder:
            self.template_dir_edit.setText(folder)
            try:
                # Persist to global JSON config
                self.config_manager.set_template_dir(folder)
                # Reinitialize template engine with new directory
                self.template_engine = EmailTemplateEngine(folder)
                self.worker.template_engine = self.template_engine
                # Reload templates and content
                self.load_templates()
                self.load_selected_template()
                InfoBar.success(
                    title="Template Updated",
                    content=f"Template folder set to: {folder}",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to set template folder: {str(e)}")

    def save_template_file(self):
        """Save current body editor content to the selected template file in the chosen folder"""
        template_name = self.template_combo.currentText()
        if template_name == "-- Select Template --" or not template_name:
            QMessageBox.warning(self, "Warning", "Please select a template to save")
            return

        template_dir = self.template_dir_edit.text().strip() or self.config_manager.get_template_dir()
        template_path = os.path.join(template_dir, template_name)

        try:
            os.makedirs(template_dir, exist_ok=True)
            content = self.email_body_edit.toPlainText()
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(content)
            InfoBar.success(
                title="Template Saved",
                content=f"Template saved: {template_path}",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )
            QMessageBox.information(self, "Success", f"Template file updated:\n{template_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save template: {str(e)}")

    def toggle_monitoring(self):
        """Toggle folder monitoring on/off"""
        if self.is_monitoring:
            self.stop_monitoring()
        else:
            self.start_monitoring()

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
                self.toggle_monitoring_btn.setText("Stop Monitoring")
                self.toggle_monitoring_btn.setIcon(FluentIcon.PAUSE)
                self.status_label.setText("Monitoring: Active")
                self.status_label.setStyleSheet("color: #107c10; font-weight: bold;")

                # Process existing files in the folder on start so not only new files are sent
                try:
                    processed_existing = self.worker.folder_monitor.process_existing_files(
                        folder_path=monitor_folder,
                        callback=self.worker.process_file,
                        key_pattern=profile_config['key_pattern'],
                        file_extensions=profile_config.get('file_extensions', [])
                    )
                    InfoBar.success(
                        title="Monitoring Started",
                        content=f"Monitoring folder: {monitor_folder} | Existing files processed: {len(processed_existing)}",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=3000,
                        parent=self
                    )
                except Exception as pe:
                    # Even if processing existing files fails, keep monitoring running
                    InfoBar.warning(
                        title="Monitoring Started",
                        content=f"Monitoring folder: {monitor_folder} | Failed to process existing files: {str(pe)}",
                        orient=Qt.Horizontal,
                        isClosable=True,
                        position=InfoBarPosition.TOP,
                        duration=5000,
                        parent=self
                    )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to start monitoring: {str(e)}")

    def stop_monitoring(self):
        """Stop folder monitoring"""
        try:
            self.worker.folder_monitor.stop_monitoring()
            self.is_monitoring = False
            self.toggle_monitoring_btn.setText("Start Monitoring")
            self.toggle_monitoring_btn.setIcon(FluentIcon.PLAY)
            self.status_label.setText("Monitoring: Stopped")
            self.status_label.setStyleSheet("color: #d73527; font-weight: bold;")
            InfoBar.success(
                title="Monitoring Stopped",
                content="Monitoring stopped",
                orient=Qt.Horizontal,
                isClosable=True,
                position=InfoBarPosition.TOP,
                duration=3000,
                parent=self
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to stop monitoring: {str(e)}")

    def save_current_config(self):
        """Save current configuration"""
        try:
            profile_name = self.profile_combo.currentData()
            if not profile_name:
                return

            # Get existing profile to preserve values not present in UI (e.g., file_extensions, subject_template)
            existing = {}
            try:
                existing = self.config_manager.get_profile_config(profile_name)
            except Exception:
                existing = {}

            # Build email_form object from UI
            selected_template = self.template_combo.currentText()
            if selected_template == "-- Select Template --":
                selected_template = existing.get('body_template', 'default_template.html')

            email_form = {
                'to_emails': [e.strip() for e in self.to_emails_edit.text().split(';') if e.strip()],
                'cc_emails': [e.strip() for e in self.cc_emails_edit.text().split(';') if e.strip()],
                'bcc_emails': [e.strip() for e in self.bcc_emails_edit.text().split(';') if e.strip()],
                'subject': self.email_subject_edit.text().strip(),
                'selected_template': selected_template
            }

            config_data = {
                # Folders and pattern
                'monitor_folder': self.monitor_folder_edit.text(),
                'sent_folder': self.sent_folder_edit.text(),
                'key_pattern': self.key_pattern_edit.text(),

                # Email client
                'email_client': self.email_client_combo.currentText(),

                # Templates
                'subject_template': existing.get('subject_template', '[filename_without_ext]'),
                'body_template': selected_template,
                
                # File extensions selected from UI (fallback to existing if none selected)
                'file_extensions': (self.get_selected_extensions() or existing.get('file_extensions', [])),
                
                # Constant variables
                'default_cc': self.default_cc_edit.text(),
                'default_bcc': self.default_bcc_edit.text(),
                'custom1_name': self.custom1_name_edit.text(),
                'custom1_value': self.custom1_value_edit.text(),
                'custom2_name': self.custom2_name_edit.text(),
                'custom2_value': self.custom2_value_edit.text(),

                # Email form fields snapshot
                'email_form': email_form
            }

            # Preserve 'name' if exists
            if 'name' in existing:
                config_data['name'] = existing['name']

            # Persist global database path and current profile
            db_path_text = self.database_path_edit.text().strip()
            if db_path_text:
                self.config_manager.set_database_path(db_path_text)
            # Persist template directory path to global config
            tpl_dir_text = self.template_dir_edit.text().strip()
            if tpl_dir_text:
                self.config_manager.set_template_dir(tpl_dir_text)
            self.config_manager.set_current_profile(profile_name)
            
            self.config_manager.save_profile_config(profile_name, config_data)

        except Exception as e:
            self.logger.error(f"Failed to save configuration: {str(e)}")

    def on_file_processed(self, file_path: str, key: str, success: bool):
        """Handle file processed signal"""
        filename = os.path.basename(file_path)
        status = "Success" if success else "Failed"

        # Add to recent files
        item = QListWidgetItem(f"{filename} ({key}) - {status}")
        # Remove FluentIcon usage for now - use simple text indicators
        if success:
            item.setText(f"✓ {filename} ({key}) - {status}")
        else:
            item.setText(f"✗ {filename} ({key}) - {status}")
        self.recent_files_list.insertItem(0, item)

        # Update files processed counter
        files_count = getattr(self, 'files_processed_count', 0) + 1
        self.files_processed_count = files_count
        self.files_processed_label.setText(f"Files Processed: {files_count}")

        # Refresh logs table
        self.refresh_logs_table()

    def on_error_occurred(self, error_message: str):
        """Handle error signal"""
        InfoBar.error(
            title="Processing Error",
            content=f"Error: {error_message}",
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP,
            duration=5000,
            parent=self
        )
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

    def load_profile_from_file(self):
        """Load profile configuration from file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Profile Configuration",
            "",
            "JSON files (*.json);;All files (*.*)"
        )

        if file_path:
            try:
                profile_name = os.path.splitext(os.path.basename(file_path))[0]
                self.config_manager.import_profile(profile_name, file_path)

                # Refresh profile combo
                self.load_current_config()

                # Set to newly imported profile
                for i in range(self.profile_combo.count()):
                    if self.profile_combo.itemData(i) == profile_name:
                        self.profile_combo.setCurrentIndex(i)
                        break

                QMessageBox.information(self, "Success", f"Profile '{profile_name}' loaded successfully!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load profile: {str(e)}")

    def save_profile_to_file(self):
        """Save profile configuration to file"""
        profile_name = self.profile_combo.currentData()
        if not profile_name:
            QMessageBox.warning(self, "Warning", "No profile selected")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Profile Configuration",
            f"{profile_name}.json",
            "JSON files (*.json);;All files (*.*)"
        )

        if file_path:
            try:
                # Save current configuration first
                self.save_current_config()

                # Export profile
                self.config_manager.export_profile(profile_name, file_path)

                QMessageBox.information(self, "Success", f"Profile '{profile_name}' saved successfully to {file_path}!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save profile: {str(e)}")

    def load_templates(self):
        """Load available templates into combo box"""
        try:
            # Prefer the path from UI if present, otherwise fallback to config
            template_dir = self.template_dir_edit.text().strip() or self.config_manager.get_template_dir()
            if not os.path.exists(template_dir):
                os.makedirs(template_dir, exist_ok=True)
            
            self.template_combo.clear()
            self.template_combo.addItem("-- Select Template --")
            
            for filename in os.listdir(template_dir):
                if filename.endswith(('.html', '.htm', '.txt')):
                    self.template_combo.addItem(filename)
            
        except Exception as e:
            self.logger.error(f"Failed to load templates: {str(e)}")

    def load_selected_template(self):
        """Load selected template content"""
        template_name = self.template_combo.currentText()
        if template_name == "-- Select Template --":
            return

        try:
            template_dir = self.template_dir_edit.text().strip() or self.config_manager.get_template_dir()
            template_path = os.path.join(template_dir, template_name)

            if os.path.exists(template_path):
                with open(template_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.email_body_edit.setPlainText(content)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load template: {str(e)}")

    def send_test_email(self):
        """Send test email using the form data"""
        try:
            # Validate email fields
            to_emails = [email.strip() for email in self.to_emails_edit.text().split(';') if email.strip()]
            if not to_emails:
                QMessageBox.warning(self, "Warning", "Please enter at least one recipient email")
                return

            cc_emails = [email.strip() for email in self.cc_emails_edit.text().split(';') if email.strip()]
            bcc_emails = [email.strip() for email in self.bcc_emails_edit.text().split(';') if email.strip()]

            subject = self.email_subject_edit.text().strip()
            if not subject:
                QMessageBox.warning(self, "Warning", "Please enter email subject")
                return

            body = self.email_body_edit.toPlainText().strip()
            if not body:
                QMessageBox.warning(self, "Warning", "Please enter email body")
                return

            # Get current profile config
            profile_config = self.config_manager.get_profile_config()

            # Create email sender
            email_sender = EmailSenderFactory.create_sender(
                profile_config['email_client'],
                **{k: v for k, v in profile_config.items() if k.startswith('smtp_')}
            )

            # Send email
            success = email_sender.send_email(
                to_emails=to_emails,
                cc_emails=cc_emails,
                bcc_emails=bcc_emails,
                subject=subject,
                body=body
            )

            if success:
                QMessageBox.information(self, "Success", "Test email sent successfully!")
            else:
                QMessageBox.critical(self, "Error", "Failed to send test email")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send test email: {str(e)}")

    def load_available_variables(self):
        """Load available variables from database and system"""
        try:
            self.variables_list.clear()

            # System variables
            system_vars = [
                "[filename] - Full filename",
                "[filename_without_ext] - Filename without extension",
                "[filepath] - Full file path",
                "[date] - Current date",
                "[time] - Current time"
            ]

            for var in system_vars:
                self.variables_list.addItem(var)

            # Database variables (from suppliers table)
            db_vars = [
                "[supplier_code] - Supplier code",
                "[supplier_name] - Supplier name",
                "[contact_name] - Contact person name",
                "[emails] - Supplier email addresses",
                "[cc_emails] - Supplier CC emails",
                "[bcc_emails] - Supplier BCC emails"
            ]

            for var in db_vars:
                self.variables_list.addItem(var)

            # Custom constant variables
            if self.custom1_name_edit.text():
                self.variables_list.addItem(f"[{self.custom1_name_edit.text()}] - {self.custom1_value_edit.text()}")
            if self.custom2_name_edit.text():
                self.variables_list.addItem(f"[{self.custom2_name_edit.text()}] - {self.custom2_value_edit.text()}")

            # Update sample data
            self.update_sample_data()

        except Exception as e:
            self.logger.error(f"Failed to load variables: {str(e)}")

    def update_sample_data(self):
        """Update sample data preview based on current variables"""
        try:
            # Sample data (this would normally come from database based on current file)
            sample_data = {
                'filename': 'TT003_invoice_2024.pdf',
                'filename_without_ext': 'TT003_invoice_2024',
                'filepath': 'C:/Monitor/TT003_invoice_2024.pdf',
                'supplier_code': 'TT003',
                'supplier_name': 'TOKO TOKO ABADI',
                'contact_name': 'John Doe',
                'emails': 'john@tokotokoabadi.com',
                'cc_emails': self.default_cc_edit.text() or 'cc@company.com',
                'bcc_emails': self.default_bcc_edit.text() or 'bcc@company.com',
                'date': '2024-01-15',
                'time': '10:30:00'
            }

            # Add custom variables
            if self.custom1_name_edit.text():
                sample_data[self.custom1_name_edit.text()] = self.custom1_value_edit.text()
            if self.custom2_name_edit.text():
                sample_data[self.custom2_name_edit.text()] = self.custom2_value_edit.text()

            # Format for display
            display_text = "Sample Data (based on TT003_invoice_2024.pdf):\n\n"
            for key, value in sample_data.items():
                display_text += f"[{key}]: {value}\n"

            self.sample_data_text.setPlainText(display_text)

        except Exception as e:
            self.logger.error(f"Failed to update sample data: {str(e)}")
            self.sample_data_text.setPlainText("Error loading sample data")

    def insert_variable_to_subject(self):
        """Insert selected variable to subject field"""
        current_item = self.variables_list.currentItem()
        if current_item:
            # Extract variable name from list item (format: "[var_name] - description")
            text = current_item.text()
            var_name = text.split('] -')[0] + ']' if '] -' in text else text

            current_subject = self.email_subject_edit.text()
            cursor_pos = self.email_subject_edit.cursorPosition()
            new_subject = current_subject[:cursor_pos] + var_name + current_subject[cursor_pos:]
            self.email_subject_edit.setText(new_subject)
            self.email_subject_edit.setCursorPosition(cursor_pos + len(var_name))

    def insert_variable_to_body(self):
        """Insert selected variable to body field"""
        current_item = self.variables_list.currentItem()
        if current_item:
            # Extract variable name from list item
            text = current_item.text()
            var_name = text.split('] -')[0] + ']' if '] -' in text else text

            cursor = self.email_body_edit.textCursor()
            cursor.insertText(var_name)

    def generate_preview(self):
        """Generate preview from email form content using real files from monitoring folder"""
        try:
            subject = self.email_subject_edit.text()
            body = self.email_body_edit.toPlainText()
            to_emails = self.to_emails_edit.text()
            cc_emails = self.cc_emails_edit.text()
            bcc_emails = self.bcc_emails_edit.text()

            # Try to get real file from monitoring folder
            monitor_folder = self.monitor_folder_edit.text().strip()
            sample_data = None
            
            if monitor_folder and os.path.exists(monitor_folder):
                # Look for files in monitoring folder that match the pattern
                key_pattern = self.key_pattern_edit.text().strip()
                if key_pattern:
                    try:
                        import re
                        pattern = re.compile(key_pattern)
                        
                        for filename in os.listdir(monitor_folder):
                            file_path = os.path.join(monitor_folder, filename)
                            if os.path.isfile(file_path):
                                match = pattern.search(filename)
                                if match:
                                    key = match.group(1) if match.groups() else match.group(0)
                                    
                                    # Try to get supplier data for this key
                                    supplier = self.database_manager.get_supplier_by_key(key)
                                    if supplier:
                                        # Use real data from file and database
                                        custom_vars = {}
                                        if self.custom1_name_edit.text():
                                            custom_vars[self.custom1_name_edit.text()] = self.custom1_value_edit.text()
                                        if self.custom2_name_edit.text():
                                            custom_vars[self.custom2_name_edit.text()] = self.custom2_value_edit.text()
                                            
                                        sample_data = self.template_engine.prepare_variables(file_path, supplier, custom_vars)
                                        sample_data['cc_emails'] = self.default_cc_edit.text() or supplier.get('cc_emails', [])
                                        sample_data['bcc_emails'] = self.default_bcc_edit.text() or supplier.get('bcc_emails', [])
                                        break
                    except Exception as e:
                        self.logger.warning(f"Error scanning monitoring folder: {str(e)}")

            # Fallback to sample data if no real files found
            if not sample_data:
                sample_data = {
                    'filename': 'TT003_invoice_2024.pdf',
                    'filename_without_ext': 'TT003_invoice_2024',
                    'filepath': os.path.join(monitor_folder or 'C:/Monitor', 'TT003_invoice_2024.pdf'),
                    'supplier_code': 'TT003',
                    'supplier_name': 'TOKO TOKO ABADI',
                    'contact_name': 'John Doe',
                    'emails': 'john@tokotokoabadi.com',
                    'cc_emails': self.default_cc_edit.text() or 'cc@company.com',
                    'bcc_emails': self.default_bcc_edit.text() or 'bcc@company.com',
                    'date': '2024-01-15',
                    'time': '10:30:00'
                }
                
                # Add custom variables
                if self.custom1_name_edit.text():
                    sample_data[self.custom1_name_edit.text()] = self.custom1_value_edit.text()
                if self.custom2_name_edit.text():
                    sample_data[self.custom2_name_edit.text()] = self.custom2_value_edit.text()

            # Process variables in subject and body
            processed_subject = self.template_engine.process_simple_variables(subject, sample_data)
            processed_body = self.template_engine.process_simple_variables(body, sample_data)

            # Create preview HTML with indication of data source
            data_source = "Real file data" if 'filepath' in sample_data and os.path.exists(sample_data['filepath']) else "Sample data"
            preview_html = f"""
            <h3>Email Preview ({data_source})</h3>
            <p><strong>To:</strong> {to_emails or '[To be filled from database]'}</p>
            <p><strong>CC:</strong> {cc_emails}</p>
            <p><strong>BCC:</strong> {bcc_emails}</p>
            <p><strong>Subject:</strong> {processed_subject}</p>
            <hr>
            <h4>Body:</h4>
            <div style="border: 1px solid #ccc; padding: 10px; background-color: #f9f9f9;">
            {processed_body.replace(chr(10), '<br>')}
            </div>
            """

            self.preview_text.setHtml(preview_html)

        except Exception as e:
            QMessageBox.warning(self, "Preview Error", f"Failed to generate preview: {str(e)}")

    def closeEvent(self, event):
        """Handle window close event"""
        if self.is_monitoring:
            self.stop_monitoring()
        event.accept()