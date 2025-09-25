import sys
import os
import time
import json
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
from PySide6.QtWidgets import QGridLayout, QVBoxLayout, QScrollArea, QWidget, QHBoxLayout

from core.database_manager import DatabaseManager  # Changed relative import
from core.config_manager import ConfigManager
from core.folder_monitor import FolderMonitor
from core.email_sender import EmailSenderFactory
from core.template_engine import EmailTemplateEngine
from utils.logger import setup_logger
from utils.resources import get_resource_path

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
            self.logger.info(f"Attempting to send email with attachment: {file_path}")
            self.logger.info(f"File exists: {os.path.exists(file_path)}")
            self.logger.info(f"File is file: {os.path.isfile(file_path)}")
            self.logger.info(f"File size: {os.path.getsize(file_path) if os.path.exists(file_path) else 'N/A'}")

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

                # Move file to sent folder AFTER successful email sending
                sent_folder = profile_config['sent_folder']
                moved_file_path = self.folder_monitor.move_file_to_sent(file_path, sent_folder)

                # Log the file move operation
                if moved_file_path:
                    self.logger.info(f"File successfully moved to sent folder: {moved_file_path}")
                else:
                    self.logger.warning(f"Failed to move file to sent folder: {file_path}")

            self.file_processed.emit(file_path, key, success)

        except Exception as e:
            self.logger.error(f"Error processing file {file_path}: {str(e)}")
            self.error_occurred.emit(f"Error processing {os.path.basename(file_path)}: {str(e)}")
        finally:
            # Ensure a 3-second delay between processing/sending files to allow file writes to complete
            try:
                time.sleep(3)
            except Exception:
                pass

class WrappingExtensionsWidget(QWidget):
    """Custom widget for displaying file extensions in a wrapping layout"""

    def __init__(self):
        super().__init__()
        self.checkboxes = {}  # Store extension -> checkbox mapping
        self.init_ui()

    def init_ui(self):
        """Initialize the wrapping layout UI"""
        layout = QVBoxLayout(self)
        layout.setSpacing(4)  # Reduced spacing
        layout.setContentsMargins(0, 0, 0, 0)

        # Create scroll area for the extensions
        self.scroll_area = QScrollArea()
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMaximumHeight(120)  # Limit height

        # Container widget for the grid layout
        self.container = QWidget()
        self.grid_layout = QGridLayout(self.container)
        self.grid_layout.setSpacing(4)  # Reduced spacing
        self.grid_layout.setContentsMargins(4, 4, 4, 4)  # Reduced margins

        self.scroll_area.setWidget(self.container)
        layout.addWidget(self.scroll_area)

    def update_extensions(self, extensions, selected):
        """Update the extensions with wrapping layout"""
        # Clear existing checkboxes
        self.clear_extensions()

        # Create checkboxes for each extension
        self.checkboxes = {}
        row = 0
        col = 0
        max_cols = 6  # Maximum columns per row - increased for better horizontal layout

        for ext in sorted(extensions):
            checkbox = CheckBox(ext)
            checkbox.setChecked(ext in selected)
            self.checkboxes[ext] = checkbox

            # Add to grid layout
            self.grid_layout.addWidget(checkbox, row, col)

            # Move to next column, wrap to next row if needed
            col += 1
            if col >= max_cols:
                col = 0
                row += 1

        # Update container size
        self.container.adjustSize()
        self.scroll_area.updateGeometry()

    def clear_extensions(self):
        """Clear all extensions"""
        # Remove all widgets from grid layout
        for i in reversed(range(self.grid_layout.count())):
            widget = self.grid_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        self.checkboxes.clear()

    def get_selected_extensions(self):
        """Get list of selected extensions"""
        return [ext for ext, checkbox in self.checkboxes.items() if checkbox.isChecked()]

    def select_all(self):
        """Select all extensions"""
        for checkbox in self.checkboxes.values():
            checkbox.setChecked(True)

    def clear_selection(self):
        """Clear all selections"""
        for checkbox in self.checkboxes.values():
            checkbox.setChecked(False)

class MainWindow(FluentWindow):
    """Main application window with Fluent Design"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Automation Desktop")
        self.resize(1200, 800)
        
        # Set window icon
        try:
            icon_path = get_resource_path('icon.ico')
            if os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
        except Exception as e:
            print(f"Failed to set window icon: {e}")
        
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
        
        self.about_interface = self.create_about_interface()
        self.about_interface.setObjectName('AboutInterface')
        
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
        
        self.addSubInterface(
            self.about_interface,
            FluentIcon.HELP,
            'About',
            NavigationItemPosition.BOTTOM
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
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Title
        title_label = TitleLabel("Email Automation Configuration")
        layout.addWidget(title_label)
        layout.addSpacing(8)

        # Profile Management Card
        profile_card = GroupHeaderCardWidget("Profile Management")
        profile_layout = QVBoxLayout()
        profile_layout.setSpacing(8)
        profile_layout.setContentsMargins(12, 12, 12, 12)

        # Profile selection row using grid layout
        profile_grid = QGridLayout()
        profile_grid.setSpacing(8)

        profile_label = StrongBodyLabel("Current Profile:")
        self.profile_combo = ComboBox()
        self.profile_combo.setMinimumWidth(200)
        self.load_profile_btn = PushButton(FluentIcon.FOLDER, "Load")
        self.save_profile_btn = PrimaryPushButton(FluentIcon.SAVE, "Save")

        profile_grid.addWidget(profile_label, 0, 0)
        profile_grid.addWidget(self.profile_combo, 0, 1)
        profile_grid.addWidget(self.load_profile_btn, 0, 2)
        profile_grid.addWidget(self.save_profile_btn, 0, 3)
        profile_grid.setColumnStretch(4, 1)

        profile_layout.addLayout(profile_grid)
        profile_card.viewLayout.addLayout(profile_layout)
        layout.addWidget(profile_card)

        # System Paths Card
        paths_card = GroupHeaderCardWidget("System Paths")
        paths_layout = QVBoxLayout()
        paths_layout.setSpacing(8)
        paths_layout.setContentsMargins(12, 12, 12, 12)

        # Database file row
        db_grid = QGridLayout()
        db_grid.setSpacing(8)
        db_label = StrongBodyLabel("Database File:")
        self.database_path_edit = LineEdit()
        self.database_path_edit.setPlaceholderText("Select database file...")
        self.database_path_edit.setMinimumWidth(300)
        self.browse_database_btn = PushButton(FluentIcon.FOLDER, "Browse")
        self.browse_database_btn.setFixedWidth(100)
        db_grid.addWidget(db_label, 0, 0)
        db_grid.addWidget(self.database_path_edit, 0, 1)
        db_grid.addWidget(self.browse_database_btn, 0, 2)
        db_grid.setColumnStretch(1, 1)
        paths_layout.addLayout(db_grid)

        # Template folder row
        tpl_grid = QGridLayout()
        tpl_grid.setSpacing(8)
        tpl_label = StrongBodyLabel("Template Folder:")
        self.template_dir_edit = LineEdit()
        self.template_dir_edit.setPlaceholderText("Select template folder...")
        self.template_dir_edit.setMinimumWidth(300)
        self.browse_template_btn = PushButton(FluentIcon.FOLDER, "Browse")
        self.browse_template_btn.setFixedWidth(100)
        tpl_grid.addWidget(tpl_label, 0, 0)
        tpl_grid.addWidget(self.template_dir_edit, 0, 1)
        tpl_grid.addWidget(self.browse_template_btn, 0, 2)
        tpl_grid.setColumnStretch(1, 1)
        paths_layout.addLayout(tpl_grid)

        paths_card.viewLayout.addLayout(paths_layout)
        layout.addWidget(paths_card)

        # Monitoring Settings Card
        monitoring_card = GroupHeaderCardWidget("Monitoring Settings")
        monitoring_layout = QVBoxLayout()
        monitoring_layout.setSpacing(8)
        monitoring_layout.setContentsMargins(12, 12, 12, 12)

        # Monitor folder row
        monitor_grid = QGridLayout()
        monitor_grid.setSpacing(8)
        monitor_label = StrongBodyLabel("Monitor Folder:")
        self.monitor_folder_edit = LineEdit()
        self.monitor_folder_edit.setPlaceholderText("Select folder to monitor...")
        self.monitor_folder_edit.setMinimumWidth(300)
        self.browse_monitor_btn = PushButton(FluentIcon.FOLDER, "Browse")
        self.browse_monitor_btn.setFixedWidth(100)
        monitor_grid.addWidget(monitor_label, 0, 0)
        monitor_grid.addWidget(self.monitor_folder_edit, 0, 1)
        monitor_grid.addWidget(self.browse_monitor_btn, 0, 2)
        monitor_grid.setColumnStretch(1, 1)
        monitoring_layout.addLayout(monitor_grid)

        # Sent folder row
        sent_grid = QGridLayout()
        sent_grid.setSpacing(8)
        sent_label = StrongBodyLabel("Sent Folder:")
        self.sent_folder_edit = LineEdit()
        self.sent_folder_edit.setPlaceholderText("Select sent files folder...")
        self.sent_folder_edit.setMinimumWidth(300)
        self.browse_sent_btn = PushButton(FluentIcon.FOLDER, "Browse")
        self.browse_sent_btn.setFixedWidth(100)
        sent_grid.addWidget(sent_label, 0, 0)
        sent_grid.addWidget(self.sent_folder_edit, 0, 1)
        sent_grid.addWidget(self.browse_sent_btn, 0, 2)
        sent_grid.setColumnStretch(1, 1)
        monitoring_layout.addLayout(sent_grid)

        # Key pattern row
        pattern_grid = QGridLayout()
        pattern_grid.setSpacing(8)
        pattern_label = StrongBodyLabel("Key Pattern (Regex):")
        self.key_pattern_edit = LineEdit()
        self.key_pattern_edit.setPlaceholderText("Enter regex pattern to extract keys from filenames...")
        self.key_pattern_edit.setMinimumWidth(300)
        pattern_grid.addWidget(pattern_label, 0, 0)
        pattern_grid.addWidget(self.key_pattern_edit, 0, 1)
        pattern_grid.setColumnStretch(1, 1)
        monitoring_layout.addLayout(pattern_grid)

        # Email client row
        client_grid = QGridLayout()
        client_grid.setSpacing(8)
        client_label = StrongBodyLabel("Email Client:")
        self.email_client_combo = ComboBox()
        self.email_client_combo.addItems(["outlook", "thunderbird", "smtp"])
        self.email_client_combo.setMinimumWidth(200)
        client_grid.addWidget(client_label, 0, 0)
        client_grid.addWidget(self.email_client_combo, 0, 1)
        client_grid.setColumnStretch(1, 1)
        monitoring_layout.addLayout(client_grid)

        monitoring_card.viewLayout.addLayout(monitoring_layout)
        layout.addWidget(monitoring_card)

        # File Extensions Card
        extensions_card = GroupHeaderCardWidget("File Types to Monitor")
        extensions_layout = QVBoxLayout()
        extensions_layout.setSpacing(8)
        extensions_layout.setContentsMargins(12, 12, 12, 12)

        # Extension controls
        ext_controls_layout = QHBoxLayout()
        ext_controls_layout.setSpacing(8)
        self.scan_extensions_btn = PushButton(FluentIcon.SEARCH, "Scan Extensions")
        self.select_all_ext_btn = PushButton(FluentIcon.CHECKBOX, "Select All")
        self.clear_ext_btn = PushButton(FluentIcon.CANCEL, "Clear")
        ext_controls_layout.addWidget(self.scan_extensions_btn)
        ext_controls_layout.addWidget(self.select_all_ext_btn)
        ext_controls_layout.addWidget(self.clear_ext_btn)
        ext_controls_layout.addStretch()
        extensions_layout.addLayout(ext_controls_layout)

        self.extensions_widget = WrappingExtensionsWidget()
        extensions_layout.addWidget(self.extensions_widget)

        extensions_card.viewLayout.addLayout(extensions_layout)
        layout.addWidget(extensions_card)

        # Variables Card
        variables_card = GroupHeaderCardWidget("Constant Variables")
        variables_layout = QVBoxLayout()
        variables_layout.setSpacing(8)
        variables_layout.setContentsMargins(12, 12, 12, 12)

        # Default CC emails row
        cc_grid = QGridLayout()
        cc_grid.setSpacing(8)
        cc_label = StrongBodyLabel("Default CC:")
        self.default_cc_edit = LineEdit()
        self.default_cc_edit.setPlaceholderText("Default CC emails (semicolon separated)")
        self.default_cc_edit.setMinimumWidth(300)
        cc_grid.addWidget(cc_label, 0, 0)
        cc_grid.addWidget(self.default_cc_edit, 0, 1)
        cc_grid.setColumnStretch(1, 1)
        variables_layout.addLayout(cc_grid)

        # Default BCC emails row
        bcc_grid = QGridLayout()
        bcc_grid.setSpacing(8)
        bcc_label = StrongBodyLabel("Default BCC:")
        self.default_bcc_edit = LineEdit()
        self.default_bcc_edit.setPlaceholderText("Default BCC emails (semicolon separated)")
        self.default_bcc_edit.setMinimumWidth(300)
        bcc_grid.addWidget(bcc_label, 0, 0)
        bcc_grid.addWidget(self.default_bcc_edit, 0, 1)
        bcc_grid.setColumnStretch(1, 1)
        variables_layout.addLayout(bcc_grid)

        # Custom variable 1 row
        custom1_grid = QGridLayout()
        custom1_grid.setSpacing(8)
        custom1_label = StrongBodyLabel("Custom Variable 1:")
        self.custom1_name_edit = LineEdit()
        self.custom1_name_edit.setPlaceholderText("Variable name")
        self.custom1_name_edit.setMinimumWidth(140)
        self.custom1_value_edit = LineEdit()
        self.custom1_value_edit.setPlaceholderText("Variable value")
        self.custom1_value_edit.setMinimumWidth(140)
        custom1_grid.addWidget(custom1_label, 0, 0)
        custom1_grid.addWidget(self.custom1_name_edit, 0, 1)
        custom1_grid.addWidget(self.custom1_value_edit, 0, 2)
        custom1_grid.setColumnStretch(1, 1)
        custom1_grid.setColumnStretch(2, 1)
        variables_layout.addLayout(custom1_grid)

        # Custom variable 2 row
        custom2_grid = QGridLayout()
        custom2_grid.setSpacing(8)
        custom2_label = StrongBodyLabel("Custom Variable 2:")
        self.custom2_name_edit = LineEdit()
        self.custom2_name_edit.setPlaceholderText("Variable name")
        self.custom2_name_edit.setMinimumWidth(140)
        self.custom2_value_edit = LineEdit()
        self.custom2_value_edit.setPlaceholderText("Variable value")
        self.custom2_value_edit.setMinimumWidth(140)
        custom2_grid.addWidget(custom2_label, 0, 0)
        custom2_grid.addWidget(self.custom2_name_edit, 0, 1)
        custom2_grid.addWidget(self.custom2_value_edit, 0, 2)
        custom2_grid.setColumnStretch(1, 1)
        custom2_grid.setColumnStretch(2, 1)
        variables_layout.addLayout(custom2_grid)

        variables_card.viewLayout.addLayout(variables_layout)
        layout.addWidget(variables_card)

        # Control Card
        control_card = SimpleCardWidget()
        control_layout = QVBoxLayout(control_card)
        control_layout.setContentsMargins(16, 16, 16, 16)

        self.toggle_monitoring_btn = PrimaryPushButton(FluentIcon.PLAY, "Start Monitoring")
        self.toggle_monitoring_btn.setFixedHeight(40)
        self.toggle_monitoring_btn.setMinimumWidth(160)
        control_layout.addWidget(self.toggle_monitoring_btn)

        layout.addWidget(control_card)
        layout.addSpacing(12)

        widget.setWidget(content_widget)
        return widget

    def create_template_interface(self):
        """Create template interface with Fluent Design"""
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(32, 32, 32, 32)

        # Title
        title_label = TitleLabel("Email Templates & Composition")
        layout.addWidget(title_label)
        layout.addSpacing(8)

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
        layout.addSpacing(12)

        # Create stacked widget for pivot content
        from PySide6.QtWidgets import QStackedWidget
        self.stackedWidget_template = QStackedWidget()
        
        # Email Form Content - wrapped in ScrollArea
        email_form_scroll = ScrollArea()
        email_form_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        email_form_scroll.setWidgetResizable(True)
        email_form_widget = self.create_email_form_content()
        email_form_scroll.setWidget(email_form_widget)
        self.stackedWidget_template.addWidget(email_form_scroll)
        
        # Variables Content - wrapped in ScrollArea
        variables_scroll = ScrollArea()
        variables_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        variables_scroll.setWidgetResizable(True)
        variables_widget = self.create_variables_content()
        variables_scroll.setWidget(variables_widget)
        self.stackedWidget_template.addWidget(variables_scroll)
        
        # Preview Content - wrapped in ScrollArea
        preview_scroll = ScrollArea()
        preview_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview_scroll.setWidgetResizable(True)
        preview_widget = self.create_preview_content()
        preview_scroll.setWidget(preview_widget)
        self.stackedWidget_template.addWidget(preview_scroll)
        
        layout.addWidget(self.stackedWidget_template, 1)  # Give it stretch factor

        return content_widget

    def create_email_form_content(self):
        """Create email form content with fluent cards"""
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(20)
        layout.setContentsMargins(0, 0, 0, 24)

        # Recipients Card
        recipients_card = GroupHeaderCardWidget("Email Recipients")
        recipients_layout = QVBoxLayout()
        recipients_layout.setSpacing(12)
        recipients_layout.setContentsMargins(24, 24, 24, 24)

        # To field
        to_row = QHBoxLayout()
        to_row.addWidget(StrongBodyLabel("To:"))
        to_row.addStretch()
        recipients_layout.addLayout(to_row)
        self.to_emails_edit = LineEdit()
        self.to_emails_edit.setPlaceholderText("Enter email addresses separated by semicolons")
        recipients_layout.addWidget(self.to_emails_edit)
        
        recipients_layout.addSpacing(8)

        # CC field
        cc_row = QHBoxLayout()
        cc_row.addWidget(StrongBodyLabel("CC:"))
        cc_row.addStretch()
        recipients_layout.addLayout(cc_row)
        self.cc_emails_edit = LineEdit()
        self.cc_emails_edit.setPlaceholderText("Enter CC email addresses separated by semicolons")
        recipients_layout.addWidget(self.cc_emails_edit)
        
        recipients_layout.addSpacing(8)

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
        content_layout.setSpacing(12)
        content_layout.setContentsMargins(24, 24, 24, 24)

        # Subject field
        subject_row = QHBoxLayout()
        subject_row.addWidget(StrongBodyLabel("Subject:"))
        subject_row.addStretch()
        content_layout.addLayout(subject_row)
        self.email_subject_edit = LineEdit()
        self.email_subject_edit.setPlaceholderText("Enter email subject...")
        content_layout.addWidget(self.email_subject_edit)
        
        content_layout.addSpacing(8)

        # Template selection
        template_row = QHBoxLayout()
        template_row.addWidget(StrongBodyLabel("Template:"))
        template_row.addStretch()
        content_layout.addLayout(template_row)
        
        template_input_row = QHBoxLayout()
        template_input_row.setSpacing(12)
        self.template_combo = ComboBox()
        self.template_combo.setMinimumWidth(300)
        self.load_templates()
        template_input_row.addWidget(self.template_combo)
        template_input_row.addStretch()
        content_layout.addLayout(template_input_row)
        
        content_layout.addSpacing(8)

        # Body field
        body_row = QHBoxLayout()
        body_row.addWidget(StrongBodyLabel("Body:"))
        body_row.addStretch()
        content_layout.addLayout(body_row)
        self.email_body_edit = TextEdit()
        self.email_body_edit.setMinimumHeight(300)
        content_layout.addWidget(self.email_body_edit)

        content_card.viewLayout.addLayout(content_layout)
        layout.addWidget(content_card)

        # Actions Card
        actions_card = SimpleCardWidget()
        actions_layout = QHBoxLayout(actions_card)
        actions_layout.setContentsMargins(24, 20, 24, 20)
        actions_layout.setSpacing(12)
        
        self.save_template_btn = PushButton(FluentIcon.SAVE, "Save Template")
        self.save_template_btn.setMinimumHeight(40)
        self.save_template_btn.setMinimumWidth(140)
        self.send_test_email_btn = PrimaryPushButton(FluentIcon.SEND, "Send Test Email")
        self.send_test_email_btn.setMinimumHeight(40)
        self.send_test_email_btn.setMinimumWidth(160)
        actions_layout.addWidget(self.save_template_btn)
        actions_layout.addWidget(self.send_test_email_btn)
        actions_layout.addStretch()

        layout.addWidget(actions_card)

        return content

    def create_variables_content(self):
        """Create variables content with fluent cards"""
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(20)
        layout.setContentsMargins(0, 0, 0, 24)

        # Available Variables Card
        variables_card = GroupHeaderCardWidget("Available Variables")
        variables_layout = QVBoxLayout()
        variables_layout.setSpacing(16)
        variables_layout.setContentsMargins(24, 24, 24, 24)

        self.variables_list = ListWidget()
        self.variables_list.setMinimumHeight(220)
        variables_layout.addWidget(self.variables_list)

        # Variable insertion buttons
        var_buttons_layout = QHBoxLayout()
        var_buttons_layout.setSpacing(12)
        self.insert_var_btn = PushButton(FluentIcon.ADD, "Insert to Subject")
        self.insert_var_btn.setMinimumHeight(40)
        self.insert_var_btn.setMinimumWidth(140)
        self.insert_var_body_btn = PushButton(FluentIcon.ADD, "Insert to Body")
        self.insert_var_body_btn.setMinimumHeight(40)
        self.insert_var_body_btn.setMinimumWidth(140)
        var_buttons_layout.addWidget(self.insert_var_btn)
        var_buttons_layout.addWidget(self.insert_var_body_btn)
        var_buttons_layout.addStretch()
        variables_layout.addLayout(var_buttons_layout)

        variables_card.viewLayout.addLayout(variables_layout)
        layout.addWidget(variables_card)

        # Sample Data Card
        sample_card = GroupHeaderCardWidget("Sample Data Preview")
        sample_layout = QVBoxLayout()
        sample_layout.setSpacing(16)
        sample_layout.setContentsMargins(24, 24, 24, 24)

        self.sample_data_text = TextEdit()
        self.sample_data_text.setReadOnly(True)
        self.sample_data_text.setMinimumHeight(200)
        self.sample_data_text.setMaximumHeight(250)
        sample_layout.addWidget(self.sample_data_text)
        
        # Load variables after sample_data_text is initialized
        self.load_available_variables()

        sample_card.viewLayout.addLayout(sample_layout)
        layout.addWidget(sample_card)

        return content

    def create_preview_content(self):
        """Create preview content with fluent cards"""
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setSpacing(20)
        layout.setContentsMargins(0, 0, 0, 24)

        # Preview Card
        preview_card = GroupHeaderCardWidget("Email Preview")
        preview_layout = QVBoxLayout()
        preview_layout.setSpacing(16)
        preview_layout.setContentsMargins(24, 24, 24, 24)

        # Preview button at top
        button_row = QHBoxLayout()
        self.preview_btn = PrimaryPushButton(FluentIcon.VIEW, "Generate Preview")
        self.preview_btn.setMinimumHeight(40)
        self.preview_btn.setMinimumWidth(180)
        button_row.addWidget(self.preview_btn)
        button_row.addStretch()
        preview_layout.addLayout(button_row)
        
        preview_layout.addSpacing(8)

        self.preview_text = TextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setMinimumHeight(400)
        preview_layout.addWidget(self.preview_text)

        preview_card.viewLayout.addLayout(preview_layout)
        layout.addWidget(preview_card)

        return content

    def create_status_interface(self):
        """Create status interface with Fluent Design"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(24)
        layout.setContentsMargins(32, 32, 32, 32)

        # Title
        title_label = TitleLabel("Status & Monitoring")
        layout.addWidget(title_label)
        layout.addSpacing(16)

        # Status Card
        status_card = GroupHeaderCardWidget("Monitoring Status")
        status_layout = QVBoxLayout()
        status_layout.setSpacing(16)
        status_layout.setContentsMargins(20, 20, 20, 20)

        # Status indicators in a grid
        status_grid = QHBoxLayout()
        status_grid.setSpacing(16)
        
        # Monitoring status indicator
        monitoring_status_card = SimpleCardWidget()
        monitoring_layout = QVBoxLayout(monitoring_status_card)
        monitoring_layout.setContentsMargins(20, 20, 20, 20)
        
        self.status_label = StrongBodyLabel("Monitoring: Stopped")
        self.status_label.setStyleSheet("color: #d73527; font-weight: bold;")
        monitoring_layout.addWidget(self.status_label)
        
        status_grid.addWidget(monitoring_status_card)
        
        # Files processed counter
        files_counter_card = SimpleCardWidget()
        files_layout = QVBoxLayout(files_counter_card)
        files_layout.setContentsMargins(20, 20, 20, 20)
        
        self.files_processed_label = StrongBodyLabel("Files Processed: 0")
        files_layout.addWidget(self.files_processed_label)
        
        status_grid.addWidget(files_counter_card)
        
        status_layout.addLayout(status_grid)
        status_card.viewLayout.addLayout(status_layout)
        layout.addWidget(status_card)

        # Recent Files Card
        recent_card = GroupHeaderCardWidget("Recent Files")
        recent_layout = QVBoxLayout()
        recent_layout.setSpacing(12)
        recent_layout.setContentsMargins(20, 20, 20, 20)

        self.recent_files_list = ListWidget()
        self.recent_files_list.setMinimumHeight(220)
        recent_layout.addWidget(self.recent_files_list)

        recent_card.viewLayout.addLayout(recent_layout)
        layout.addWidget(recent_card)

        # Email Logs Card
        logs_card = GroupHeaderCardWidget("Email Logs")
        logs_layout = QVBoxLayout()
        logs_layout.setSpacing(12)
        logs_layout.setContentsMargins(20, 20, 20, 20)

        self.logs_table = TableWidget()
        self.logs_table.setColumnCount(7)
        self.logs_table.setHorizontalHeaderLabels([
            "Time", "File", "Supplier", "Subject", "Recipients", "Email Client", "Status"
        ])

        # Set column widths for better visibility
        self.logs_table.setColumnWidth(0, 140)  # Time - wider for full timestamp
        self.logs_table.setColumnWidth(1, 200)  # File - wider for full filename
        self.logs_table.setColumnWidth(2, 120)  # Supplier - for supplier info
        self.logs_table.setColumnWidth(3, 250)  # Subject - for email subject
        self.logs_table.setColumnWidth(4, 200)  # Recipients - for email addresses
        self.logs_table.setColumnWidth(5, 100)  # Email Client - for client type
        self.logs_table.setColumnWidth(6, 80)   # Status - for status info

        # Make columns resizable
        header = self.logs_table.horizontalHeader()
        header.setStretchLastSection(False)  # Don't stretch last column
        header.setSectionResizeMode(0, header.ResizeMode.Fixed)      # Time
        header.setSectionResizeMode(1, header.ResizeMode.Interactive) # File
        header.setSectionResizeMode(2, header.ResizeMode.Interactive) # Supplier
        header.setSectionResizeMode(3, header.ResizeMode.Interactive) # Subject
        header.setSectionResizeMode(4, header.ResizeMode.Interactive) # Recipients
        header.setSectionResizeMode(5, header.ResizeMode.Fixed)       # Email Client
        header.setSectionResizeMode(6, header.ResizeMode.Fixed)       # Status

        # Enable text wrapping for long content
        self.logs_table.setWordWrap(True)
        self.logs_table.setTextElideMode(Qt.TextElideMode.ElideNone)

        # Set minimum height for better visibility
        self.logs_table.setMinimumHeight(400)
        self.logs_table.setMaximumHeight(600)

        # Enable horizontal scrolling if needed
        self.logs_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.logs_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Enable sorting
        self.logs_table.setSortingEnabled(True)
        self.logs_table.sortByColumn(0, Qt.SortOrder.DescendingOrder)  # Sort by time descending

        # Set selection behavior
        self.logs_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.logs_table.setAlternatingRowColors(True)

        logs_layout.addWidget(self.logs_table)

        logs_card.viewLayout.addLayout(logs_layout)
        layout.addWidget(logs_card)

        layout.addSpacing(24)
        widget.setWidget(content_widget)
        return widget

    def create_about_interface(self):
        """Create about interface with usage guide and developer info"""
        widget = ScrollArea()
        widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        widget.setWidgetResizable(True)
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setSpacing(24)
        layout.setContentsMargins(32, 32, 32, 32)

        # Title
        title_label = TitleLabel("About Email Automation Desktop")
        layout.addWidget(title_label)
        layout.addSpacing(16)

        # Application Info Card
        app_info_card = GroupHeaderCardWidget("Application Information")
        app_info_layout = QVBoxLayout()
        app_info_layout.setSpacing(12)
        app_info_layout.setContentsMargins(20, 20, 20, 20)

        app_info_text = BodyLabel(
            "Email Automation Desktop v1.0.0\n"
            "Automated email sending based on file monitoring\n"
            "Built with PySide6 and PyQt-Fluent-Widgets"
        )
        app_info_layout.addWidget(app_info_text)

        app_info_card.viewLayout.addLayout(app_info_layout)
        layout.addWidget(app_info_card)

        # Quick Usage Guide Card
        usage_card = GroupHeaderCardWidget("Quick Usage Guide")
        usage_layout = QVBoxLayout()
        usage_layout.setSpacing(16)
        usage_layout.setContentsMargins(20, 20, 20, 20)

        usage_steps = [
            "1. Configure Database & Templates",
            "   • Set database file path (suppliers data)",
            "   • Set template folder for email templates",
            "",
            "2. Setup Monitoring",
            "   • Choose folder to monitor for new files",
            "   • Set sent folder for processed files",
            "   • Define key pattern (regex) to extract supplier codes",
            "   • Select file extensions to monitor (.pdf, .xlsx, etc.)",
            "",
            "3. Configure Email Settings",
            "   • Choose email client (Outlook/Thunderbird/SMTP)",
            "   • Set default CC/BCC emails if needed",
            "   • Add custom variables for templates",
            "",
            "4. Design Email Templates",
            "   • Create email templates with variables",
            "   • Use Variables tab to see available placeholders",
            "   • Preview emails before sending",
            "",
            "5. Start Monitoring",
            "   • Click 'Start Monitoring' button",
            "   • Application will automatically process new files",
            "   • Check Status & Logs for monitoring activity",
            "",
            "6. Monitor Results",
            "   • View recent processed files",
            "   • Check email logs for sent emails",
            "   • Monitor processing status"
        ]

        usage_text = BodyLabel('\n'.join(usage_steps))
        usage_text.setWordWrap(True)
        usage_layout.addWidget(usage_text)

        usage_card.viewLayout.addLayout(usage_layout)
        layout.addWidget(usage_card)

        # Variable Examples Card
        variables_card = GroupHeaderCardWidget("Template Variables Examples")
        variables_layout = QVBoxLayout()
        variables_layout.setSpacing(12)
        variables_layout.setContentsMargins(20, 20, 20, 20)

        variables_examples = [
            "System Variables:",
            "• [filename] - Full filename with extension",
            "• [filename_without_ext] - Filename without extension",
            "• [filepath] - Complete file path",
            "• [date] - Current date",
            "• [time] - Current time",
            "",
            "Supplier Variables (from database):",
            "• [supplier_code] - Supplier identification code",
            "• [supplier_name] - Company/supplier name",
            "• [contact_name] - Contact person name",
            "• [emails] - Primary email addresses",
            "",
            "Example Usage in Templates:",
            'Subject: "Invoice [filename_without_ext] for [supplier_name]"',
            'Body: "Dear [contact_name], please find attached [filename]..."'
        ]

        variables_text = BodyLabel('\n'.join(variables_examples))
        variables_text.setWordWrap(True)
        variables_layout.addWidget(variables_text)

        variables_card.viewLayout.addLayout(variables_layout)
        layout.addWidget(variables_card)

        # Troubleshooting Card
        troubleshooting_card = GroupHeaderCardWidget("Troubleshooting Tips")
        troubleshooting_layout = QVBoxLayout()
        troubleshooting_layout.setSpacing(12)
        troubleshooting_layout.setContentsMargins(20, 20, 20, 20)

        troubleshooting_tips = [
            "Common Issues & Solutions:",
            "",
            "• Files not being processed:",
            "  - Check monitor folder path exists",
            "  - Verify key pattern regex is correct",
            "  - Ensure file extensions are selected",
            "  - Check supplier exists in database for extracted key",
            "",
            "• Emails not sending:",
            "  - Verify email client is configured properly",
            "  - Check internet connection",
            "  - Ensure Outlook is running (for Outlook client)",
            "  - Verify SMTP settings (for Thunderbird/SMTP)",
            "",
            "• Template variables not working:",
            "  - Use correct variable syntax: [variable_name]",
            "  - Check Variables tab for available options",
            "  - Ensure supplier data exists in database",
            "",
            "• Performance issues:",
            "  - Limit file extensions to necessary types only",
            "  - Keep monitoring folder organized",
            "  - Regular database maintenance"
        ]

        troubleshooting_text = BodyLabel('\n'.join(troubleshooting_tips))
        troubleshooting_text.setWordWrap(True)
        troubleshooting_layout.addWidget(troubleshooting_text)

        troubleshooting_card.viewLayout.addLayout(troubleshooting_layout)
        layout.addWidget(troubleshooting_card)

        # Developer Credit Card
        credit_card = SimpleCardWidget()
        credit_layout = QVBoxLayout(credit_card)
        credit_layout.setContentsMargins(24, 20, 24, 20)
        credit_layout.setSpacing(8)

        # Center the developer info
        credit_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        developed_label = StrongBodyLabel("Developed by")
        developed_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_layout.addWidget(developed_label)
        
        developer_label = TitleLabel("Faizal Kusmawan")
        developer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        developer_label.setStyleSheet("color: #0078d4; font-weight: bold;")
        credit_layout.addWidget(developer_label)
        
        # Add some spacing
        credit_layout.addSpacing(8)
        
        version_label = CaptionLabel("Email Automation Desktop v1.0.0")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        credit_layout.addWidget(version_label)

        layout.addWidget(credit_card)
        layout.addSpacing(24)

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
        self.profile_names = []  # Store profile names separately
        
        for profile in profiles:
            self.profile_combo.addItem(profile['display_name'])
            self.profile_names.append(profile['name'])

        # Set current profile
        current_profile = self.config_manager.get_current_profile()
        for i, profile_name in enumerate(self.profile_names):
            if profile_name == current_profile:
                self.profile_combo.setCurrentIndex(i)
                break

        self.load_profile_config()

    def load_profile_config(self):
        """Load configuration for selected profile"""
        current_index = self.profile_combo.currentIndex()
        if current_index < 0 or current_index >= len(self.profile_names):
            return
        
        profile_name = self.profile_names[current_index]

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
            # For fluent ComboBox, we need to find the text manually
            for i in range(self.email_client_combo.count()):
                if self.email_client_combo.itemText(i) == client:
                    self.email_client_combo.setCurrentIndex(i)
                    break

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
                # Search for template manually in fluent ComboBox
                template_found = False
                for i in range(self.template_combo.count()):
                    if self.template_combo.itemText(i) == selected_template:
                        self.template_combo.setCurrentIndex(i)
                        template_found = True
                        break
                
                if not template_found:
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
            self.extensions_widget.update_extensions(extensions, selected)
        except Exception as e:
            self.logger.error(f"Failed to update extensions list: {str(e)}")

    def get_selected_extensions(self):
        """Get currently selected extensions from the widget"""
        try:
            return self.extensions_widget.get_selected_extensions()
        except Exception as e:
            self.logger.error(f"Failed to read selected extensions: {str(e)}")
            return []

    def select_all_extensions(self):
        """Select all extensions in the widget"""
        try:
            self.extensions_widget.select_all()
        except Exception as e:
            self.logger.error(f"Failed to select all extensions: {str(e)}")

    def clear_extensions_selection(self):
        """Clear all selections in the extensions widget"""
        try:
            self.extensions_widget.clear_selection()
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
            current_index = self.profile_combo.currentIndex()
            if current_index < 0 or current_index >= len(self.profile_names):
                return
            
            profile_name = self.profile_names[current_index]

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
        """Refresh email logs table with enhanced visibility"""
        try:
            logs = self.database_manager.get_email_logs(limit=100)
            self.logs_table.setRowCount(len(logs))

            for row, log in enumerate(logs):
                # Time column - format timestamp for better readability
                sent_at = log.get('sent_at', '')
                if sent_at:
                    # Try to format the timestamp if it's in ISO format
                    try:
                        from datetime import datetime
                        dt = datetime.fromisoformat(sent_at.replace('Z', '+00:00'))
                        formatted_time = dt.strftime('%Y-%m-%d\n%H:%M:%S')
                    except:
                        formatted_time = sent_at
                else:
                    formatted_time = ''
                self.logs_table.setItem(row, 0, QTableWidgetItem(formatted_time))

                # File column - show full filename
                filename = log.get('filename', '')
                # Show full filename, add tooltip with full path if available
                file_path = log.get('file_path', '')
                filename_item = QTableWidgetItem(filename)
                if file_path and file_path != filename:
                    filename_item.setToolTip(f"Full path: {file_path}")
                self.logs_table.setItem(row, 1, filename_item)

                # Supplier column - show both key and name
                supplier_key = log.get('supplier_key', '')
                supplier_info = supplier_key

                # Try to get supplier name from database
                try:
                    supplier = self.database_manager.get_supplier_by_key(supplier_key)
                    if supplier:
                        supplier_name = supplier.get('supplier_name', '')
                        if supplier_name and supplier_name != supplier_key:
                            supplier_info = f"{supplier_key}\n{supplier_name}"
                except:
                    pass  # Keep original supplier_key if lookup fails

                supplier_item = QTableWidgetItem(supplier_info)
                # Add tooltip with additional supplier information
                try:
                    supplier = self.database_manager.get_supplier_by_key(supplier_key)
                    if supplier:
                        contact_name = supplier.get('contact_name', '')
                        emails = supplier.get('emails', [])
                        cc_emails = supplier.get('cc_emails', [])
                        bcc_emails = supplier.get('bcc_emails', [])

                        tooltip_parts = []
                        if contact_name:
                            tooltip_parts.append(f"Contact: {contact_name}")
                        if emails:
                            tooltip_parts.append(f"Emails: {', '.join(emails)}")
                        if cc_emails:
                            tooltip_parts.append(f"CC: {', '.join(cc_emails)}")
                        if bcc_emails:
                            tooltip_parts.append(f"BCC: {', '.join(bcc_emails)}")

                        if tooltip_parts:
                            supplier_item.setToolTip('\n'.join(tooltip_parts))
                except:
                    pass  # Keep simple tooltip if lookup fails

                self.logs_table.setItem(row, 2, supplier_item)

                # Subject column - show email subject
                subject = log.get('subject', '')
                # Truncate very long subjects for display
                display_subject = subject
                if len(subject) > 60:
                    display_subject = subject[:57] + "..."

                subject_item = QTableWidgetItem(display_subject)
                # Add full subject as tooltip if truncated
                if len(subject) > 60:
                    subject_item.setToolTip(subject)
                self.logs_table.setItem(row, 3, subject_item)

                # Recipients column - show recipient emails
                recipient_emails = log.get('recipient_emails', '[]')
                try:
                    # Handle empty, null, or invalid JSON data
                    if not recipient_emails or recipient_emails == '[]' or recipient_emails == '':
                        emails = []
                    else:
                        emails = json.loads(recipient_emails)

                    if emails and isinstance(emails, list):
                        recipients_text = '\n'.join(emails[:3])  # Show first 3 emails
                        if len(emails) > 3:
                            recipients_text += f"\n... (+{len(emails) - 3} more)"

                        # Add full list as tooltip
                        full_recipients = '\n'.join(emails)
                        recipients_item = QTableWidgetItem(recipients_text)
                        recipients_item.setToolTip(f"All recipients:\n{full_recipients}")
                    else:
                        recipients_text = 'No recipients'
                        recipients_item = QTableWidgetItem(recipients_text)
                except (json.JSONDecodeError, TypeError) as e:
                    # Handle malformed JSON or other parsing errors
                    if recipient_emails and recipient_emails != '[]':
                        # Try to extract emails from malformed JSON or show raw data
                        recipients_text = str(recipient_emails)[:50] + "..." if len(str(recipient_emails)) > 50 else str(recipient_emails)
                        recipients_item = QTableWidgetItem(recipients_text)
                        recipients_item.setToolTip(f"Raw data: {recipient_emails}")
                    else:
                        recipients_text = 'No recipients'
                        recipients_item = QTableWidgetItem(recipients_text)

                self.logs_table.setItem(row, 4, recipients_item)

                # Email Client column
                email_client = log.get('email_client', '')
                client_item = QTableWidgetItem(email_client)
                self.logs_table.setItem(row, 5, client_item)

                # Status column - enhanced status display
                status = log.get('status', 'sent')
                status_item = QTableWidgetItem(status.upper())

                # Color code status
                if status.lower() == 'sent':
                    status_item.setBackground(Qt.GlobalColor.green)
                    status_item.setForeground(Qt.GlobalColor.white)
                elif status.lower() == 'failed':
                    status_item.setBackground(Qt.GlobalColor.red)
                    status_item.setForeground(Qt.GlobalColor.white)
                elif status.lower() == 'pending':
                    status_item.setBackground(Qt.GlobalColor.yellow)
                    status_item.setForeground(Qt.GlobalColor.black)

                self.logs_table.setItem(row, 6, status_item)

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
                for i, stored_profile_name in enumerate(self.profile_names):
                    if stored_profile_name == profile_name:
                        self.profile_combo.setCurrentIndex(i)
                        break

                QMessageBox.information(self, "Success", f"Profile '{profile_name}' loaded successfully!")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load profile: {str(e)}")

    def save_profile_to_file(self):
        """Save profile configuration to file"""
        current_index = self.profile_combo.currentIndex()
        if current_index < 0 or current_index >= len(self.profile_names):
            QMessageBox.warning(self, "Warning", "No profile selected")
            return
            
        profile_name = self.profile_names[current_index]

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