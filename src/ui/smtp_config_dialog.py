import sys
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                               QLineEdit, QSpinBox, QCheckBox, QPushButton,
                               QLabel, QMessageBox)
from PySide6.QtCore import Qt
from qfluentwidgets import (PrimaryPushButton, PushButton, BodyLabel, TitleLabel,
                           InfoBar, InfoBarPosition, CardWidget, SimpleCardWidget, ComboBox)

from core.email_sender import ThunderbirdSender


class SMTPConfigDialog(QDialog):
    """Dialog for configuring SMTP settings for Thunderbird/SMTP email client"""

    def __init__(self, parent=None, smtp_config=None):
        super().__init__(parent)
        self.smtp_config = smtp_config or {}
        self.test_sender = None

        self.setWindowTitle("SMTP Configuration")
        self.setModal(True)
        self.resize(500, 400)

        self.init_ui()
        self.load_config()

    def init_ui(self):
        """Initialize the dialog UI"""
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Title
        title_label = TitleLabel("SMTP Settings")
        layout.addWidget(title_label)

        # Configuration card
        config_card = CardWidget()
        config_layout = QVBoxLayout(config_card)
        config_layout.setContentsMargins(20, 20, 20, 20)

        # Form layout for SMTP fields
        form_layout = QFormLayout()
        form_layout.setSpacing(12)

        # Server field
        self.server_edit = QLineEdit()
        self.server_edit.setPlaceholderText("e.g., smtp.gmail.com")
        form_layout.addRow("SMTP Server:", self.server_edit)

        # Port field with dropdown
        self.port_combo = ComboBox()
        self.port_combo.addItems([
            "587 (TLS)",
            "465 (SSL)",
            "25 (Plain)",
            "Custom"
        ])
        self.port_combo.setCurrentText("587 (TLS)")
        self.port_combo.currentTextChanged.connect(self.on_port_changed)

        # Custom port spin box (initially hidden)
        self.custom_port_spin = QSpinBox()
        self.custom_port_spin.setRange(1, 65535)
        self.custom_port_spin.setValue(587)
        self.custom_port_spin.setVisible(False)

        port_layout = QHBoxLayout()
        port_layout.addWidget(self.port_combo)
        port_layout.addWidget(self.custom_port_spin)
        port_layout.addStretch()

        form_layout.addRow("Port:", port_layout)

        # Username field
        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("your_email@gmail.com")
        form_layout.addRow("Username:", self.username_edit)

        # Password field
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setPlaceholderText("Your email password or app password")
        form_layout.addRow("Password:", self.password_edit)

        # TLS checkbox
        self.tls_checkbox = QCheckBox("Use TLS (recommended for port 587)")
        self.tls_checkbox.setChecked(True)
        form_layout.addRow("", self.tls_checkbox)

        config_layout.addLayout(form_layout)

        # Test connection button
        test_layout = QHBoxLayout()
        test_layout.addStretch()
        self.test_connection_btn = PushButton("Test Connection")
        self.test_connection_btn.clicked.connect(self.test_connection)
        test_layout.addWidget(self.test_connection_btn)
        config_layout.addLayout(test_layout)

        layout.addWidget(config_card)

        # Help text
        help_card = SimpleCardWidget()
        help_layout = QVBoxLayout(help_card)
        help_layout.setContentsMargins(16, 16, 16, 16)

        help_text = BodyLabel(
            "For Gmail: Use smtp.gmail.com with port 587, enable TLS, and use an App Password.\n"
            "For Outlook.com: Use smtp-mail.outlook.com with port 587 and TLS.\n"
            "For Yahoo: Use smtp.mail.yahoo.com with port 587 and TLS."
        )
        help_text.setWordWrap(True)
        help_layout.addWidget(help_text)

        layout.addWidget(help_card)

        # Dialog buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()

        self.cancel_btn = PushButton("Cancel")
        self.cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(self.cancel_btn)

        self.save_btn = PrimaryPushButton("Save")
        self.save_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(self.save_btn)

        layout.addLayout(buttons_layout)

    def load_config(self):
        """Load SMTP configuration into the dialog"""
        self.server_edit.setText(self.smtp_config.get('smtp_server', ''))

        # Set port based on configuration
        port = int(self.smtp_config.get('smtp_port', 587))
        if port == 587:
            self.port_combo.setCurrentText("587 (TLS)")
        elif port == 465:
            self.port_combo.setCurrentText("465 (SSL)")
        elif port == 25:
            self.port_combo.setCurrentText("25 (Plain)")
        else:
            self.port_combo.setCurrentText("Custom")
            self.custom_port_spin.setValue(port)
            self.custom_port_spin.setVisible(True)

        self.username_edit.setText(self.smtp_config.get('smtp_username', ''))
        self.password_edit.setText(self.smtp_config.get('smtp_password', ''))
        self.tls_checkbox.setChecked(self.smtp_config.get('smtp_use_tls', True))

    def get_config(self):
        """Get SMTP configuration from the dialog"""
        port_text = self.port_combo.currentText()
        if port_text == "587 (TLS)":
            port = 587
        elif port_text == "465 (SSL)":
            port = 465
        elif port_text == "25 (Plain)":
            port = 25
        else:  # Custom
            port = self.custom_port_spin.value()

        return {
            'smtp_server': self.server_edit.text().strip(),
            'smtp_port': port,
            'smtp_username': self.username_edit.text().strip(),
            'smtp_password': self.password_edit.text(),
            'smtp_use_tls': self.tls_checkbox.isChecked()
        }

    def on_port_changed(self):
        """Handle port combo box change"""
        port_text = self.port_combo.currentText()
        if port_text == "Custom":
            self.custom_port_spin.setVisible(True)
        else:
            self.custom_port_spin.setVisible(False)
            # Set the spin box value to match the selected preset
            if port_text == "587 (TLS)":
                self.custom_port_spin.setValue(587)
            elif port_text == "465 (SSL)":
                self.custom_port_spin.setValue(465)
            elif port_text == "25 (Plain)":
                self.custom_port_spin.setValue(25)

    def test_connection(self):
        """Test SMTP connection with current settings"""
        config = self.get_config()

        # Validate required fields
        if not config['smtp_server']:
            QMessageBox.warning(self, "Validation Error", "SMTP Server is required")
            return
        if not config['smtp_username']:
            QMessageBox.warning(self, "Validation Error", "Username is required")
            return
        if not config['smtp_password']:
            QMessageBox.warning(self, "Validation Error", "Password is required")
            return

        # Disable test button during test
        self.test_connection_btn.setEnabled(False)
        self.test_connection_btn.setText("Testing...")

        try:
            # Create test sender
            self.test_sender = ThunderbirdSender(
                smtp_server=config['smtp_server'],
                smtp_port=config['smtp_port'],
                username=config['smtp_username'],
                password=config['smtp_password'],
                use_tls=config['smtp_use_tls']
            )

            # Test connection
            if self.test_sender.test_connection():
                InfoBar.success(
                    title="Connection Successful",
                    content="SMTP connection test passed!",
                    orient=Qt.Horizontal,
                    isClosable=True,
                    position=InfoBarPosition.TOP,
                    duration=3000,
                    parent=self
                )
                QMessageBox.information(self, "Success", "SMTP connection test successful!")
            else:
                QMessageBox.warning(self, "Connection Failed", "SMTP connection test failed. Please check your settings.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Connection test failed: {str(e)}")

        finally:
            # Re-enable test button
            self.test_connection_btn.setEnabled(True)
            self.test_connection_btn.setText("Test Connection")

    def accept(self):
        """Validate and accept the dialog"""
        config = self.get_config()

        # Validate required fields
        if not config['smtp_server']:
            QMessageBox.warning(self, "Validation Error", "SMTP Server is required")
            return
        if not config['smtp_username']:
            QMessageBox.warning(self, "Validation Error", "Username is required")
            return
        if not config['smtp_password']:
            QMessageBox.warning(self, "Validation Error", "Password is required")
            return

        super().accept()