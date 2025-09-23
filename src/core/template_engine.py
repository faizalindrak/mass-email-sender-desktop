from jinja2 import Template, Environment, FileSystemLoader
import os
import re
import datetime
from typing import Dict, Any, Optional

class EmailTemplateEngine:
    """Template engine for email content generation"""

    def __init__(self, template_dir: str = None):
        self.template_dir = template_dir or "templates"
        if os.path.exists(self.template_dir):
            self.env = Environment(loader=FileSystemLoader(self.template_dir))
        else:
            self.env = Environment()

    def render_template(self, template_content: str, variables: Dict[str, Any]) -> str:
        """Render template with variables"""
        try:
            template = Template(template_content)
            return template.render(**variables)
        except Exception as e:
            raise Exception(f"Template rendering error: {str(e)}")

    def render_file_template(self, template_filename: str, variables: Dict[str, Any]) -> str:
        """Render template from file"""
        try:
            template = self.env.get_template(template_filename)
            return template.render(**variables)
        except Exception as e:
            raise Exception(f"Template file rendering error: {str(e)}")

    def prepare_variables(self, file_path: str, supplier_data: Dict[str, Any], custom_vars: Dict[str, Any] = None) -> Dict[str, Any]:
        """Prepare variables for template rendering"""
        filename = os.path.basename(file_path)
        filename_without_ext = os.path.splitext(filename)[0]
        file_size = os.path.getsize(file_path) if os.path.exists(file_path) else 0

        now = datetime.datetime.now()

        variables = {
            # File variables
            'filename': filename,
            'filename_without_ext': filename_without_ext,
            'filepath': file_path,
            'file_size': file_size,
            'file_size_mb': round(file_size / (1024 * 1024), 2),

            # Supplier variables
            'supplier_code': supplier_data.get('supplier_code', ''),
            'supplier_name': supplier_data.get('supplier_name', ''),
            'contact_name': supplier_data.get('contact_name', ''),
            'emails': supplier_data.get('emails', []),
            'cc_emails': supplier_data.get('cc_emails', []),
            'bcc_emails': supplier_data.get('bcc_emails', []),

            # Date/time variables
            'date': now.strftime('%Y-%m-%d'),
            'time': now.strftime('%H:%M:%S'),
            'datetime': now.strftime('%Y-%m-%d %H:%M:%S'),
            'date_indonesian': now.strftime('%d %B %Y'),
            'day': now.strftime('%d'),
            'month': now.strftime('%m'),
            'year': now.strftime('%Y'),
            'month_name': now.strftime('%B'),
            'day_name': now.strftime('%A'),

            # System variables
            'current_user': os.getenv('USERNAME', 'User'),
            'computer_name': os.getenv('COMPUTERNAME', 'Computer'),
        }

        # Add custom variables if provided
        if custom_vars:
            variables.update(custom_vars)

        return variables

    def process_simple_variables(self, text: str, variables: Dict[str, Any]) -> str:
        """Process simple variables in [variable_name] format"""
        def replace_var(match):
            var_name = match.group(1)
            value = variables.get(var_name, f'[{var_name}]')

            # Handle list variables
            if isinstance(value, list):
                return ', '.join(str(v) for v in value)

            return str(value)

        return re.sub(r'\\[(\\w+)\\]', replace_var, text)

    def validate_template(self, template_content: str) -> tuple[bool, str]:
        """Validate template syntax"""
        try:
            Template(template_content)
            return True, "Template is valid"
        except Exception as e:
            return False, f"Template validation error: {str(e)}"

    def get_template_variables(self, template_content: str) -> list:
        """Extract variables from template"""
        try:
            template = Template(template_content)
            ast = template.environment.parse(template_content)
            variables = set()

            for node in ast.find_all():
                if hasattr(node, 'name'):
                    variables.add(node.name)

            return list(variables)
        except:
            return []

    def create_default_templates(self):
        """Create default template files"""
        os.makedirs(self.template_dir, exist_ok=True)

        # Default invoice template
        invoice_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #f0f0f0; padding: 15px; border-radius: 5px; }
        .content { margin: 20px 0; }
        .footer { font-size: 12px; color: #666; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Invoice Order - {{ filename_without_ext }}</h2>
        <p><strong>Supplier:</strong> {{ supplier_name }} ({{ supplier_code }})</p>
        <p><strong>Contact:</strong> {{ contact_name }}</p>
        <p><strong>Date:</strong> {{ date_indonesian }}</p>
    </div>

    <div class="content">
        <p>Dear {{ contact_name }},</p>
        <p>Terlampir invoice order dengan detail sebagai berikut:</p>
        <ul>
            <li>File: {{ filename }}</li>
            <li>Size: {{ file_size_mb }} MB</li>
            <li>Date: {{ datetime }}</li>
        </ul>
        <p>Mohon untuk review dan proses sesuai prosedur yang berlaku.</p>
    </div>

    <div class="footer">
        <p>Email ini dikirim secara otomatis dari sistem Email Automation.</p>
        <p>Generated on {{ datetime }} by {{ current_user }}</p>
    </div>
</body>
</html>
        """

        # Default delivery template
        delivery_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #e8f5e8; padding: 15px; border-radius: 5px; }
        .content { margin: 20px 0; }
        .footer { font-size: 12px; color: #666; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Delivery Schedule - {{ filename_without_ext }}</h2>
        <p><strong>Supplier:</strong> {{ supplier_name }} ({{ supplier_code }})</p>
        <p><strong>Contact:</strong> {{ contact_name }}</p>
        <p><strong>Date:</strong> {{ date_indonesian }}</p>
    </div>

    <div class="content">
        <p>Dear {{ contact_name }},</p>
        <p>Terlampir schedule delivery dengan detail sebagai berikut:</p>
        <ul>
            <li>File: {{ filename }}</li>
            <li>Schedule Date: {{ date }}</li>
            <li>Generated: {{ datetime }}</li>
        </ul>
        <p>Mohon untuk mempersiapkan delivery sesuai schedule yang telah ditentukan.</p>
    </div>

    <div class="footer">
        <p>Email ini dikirim secara otomatis dari sistem Email Automation.</p>
        <p>Generated on {{ datetime }} by {{ current_user }}</p>
    </div>
</body>
</html>
        """

        # Simple default template
        default_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .content { margin: 20px 0; }
    </style>
</head>
<body>
    <div class="content">
        <p>Dear {{ contact_name }},</p>
        <p>Terlampir file {{ filename }} untuk {{ supplier_name }}.</p>
        <p>Mohon untuk review dan proses sesuai kebutuhan.</p>
        <br>
        <p>Best regards,<br>{{ current_user }}</p>
    </div>
</body>
</html>
        """

        templates = {
            'invoice_template.html': invoice_template,
            'delivery_template.html': delivery_template,
            'default_template.html': default_template
        }

        for filename, content in templates.items():
            filepath = os.path.join(self.template_dir, filename)
            if not os.path.exists(filepath):
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(content.strip())

    def preview_template(self, template_content: str, sample_data: Dict[str, Any] = None) -> str:
        """Generate preview of template with sample data"""
        if not sample_data:
            sample_data = {
                'filename': 'INV001_TT003_20241201.pdf',
                'filename_without_ext': 'INV001_TT003_20241201',
                'filepath': 'C:\\Orders\\INV001_TT003_20241201.pdf',
                'supplier_code': 'TT003',
                'supplier_name': 'TOKO TOKO ABADI',
                'contact_name': 'Budi Santoso',
                'emails': ['budi@tokoabadi.com'],
                'cc_emails': ['manager@tokoabadi.com'],
                'date': '2024-12-01',
                'time': '14:30:00',
                'current_user': 'Admin'
            }

        return self.render_template(template_content, sample_data)