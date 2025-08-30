from fpdf import FPDF
import os
from config import Config

class CertificatePDF(FPDF):
    def __init__(self, organization):
        super().__init__()
        self.organization = organization
    
    def header(self):
        self.set_font('Helvetica', 'B', 24)
        self.cell(0, 10, 'Certificate of Participation', ln=True, align='C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', align='C')

class CertificateGenerator:
    def __init__(self):
        self.config = Config()
    
    def generate_certificate(self, name, roll_number, event_name, event_date):
        """Generate a PDF certificate for a participant"""
        pdf = CertificatePDF(self.config.ORGANIZATION)
        pdf.add_page()
        
        # Title
        pdf.set_font('Times', 'B', 24)
        pdf.set_text_color(102, 126, 234)
        pdf.cell(0, 15, 'Certificate of Participation', ln=True, align='C')

        # Subtitle
        pdf.set_font('Helvetica', '', 12)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(0, 10, 'This is to certify that', ln=True, align='C')

        # Awarded to
        pdf.set_font('Helvetica', '', 14)
        pdf.set_text_color(71, 85, 105)
        pdf.cell(0, 10, 'This certificate is proudly awarded to', ln=True, align='C')

        # Participant name
        pdf.set_font('Times', 'B', 22)
        pdf.set_text_color(30, 41, 59)
        pdf.cell(0, 15, name, ln=True, align='C')
        pdf.ln(10)

        # Achievement text
        pdf.set_font('Helvetica', '', 12)
        pdf.set_text_color(71, 85, 105)
        pdf.multi_cell(0, 10, f'For active participation and demonstrating exceptional enthusiasm in the', align='C')
        pdf.ln(10)

        # Event details box
        pdf.set_fill_color(102, 126, 234)
        pdf.set_draw_color(102, 126, 234)
        pdf.set_line_width(0.5)
        x = pdf.get_x()
        y = pdf.get_y()
        width = pdf.w - 2 * pdf.l_margin
        height = 25
        pdf.rect(x, y, width, height, style='DF')
        pdf.set_xy(x, y + 5)
        pdf.set_font('Times', 'B', 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, event_name, ln=True, align='C')
        pdf.set_font('Helvetica', '', 10)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 10, f'Organized on {event_date}', ln=True, align='C')
        pdf.ln(15)

        # Signatures
        pdf.set_font('Times', 'B', 14)
        pdf.set_text_color(102, 126, 234)
        pdf.cell(90, 10, 'Principal Revathi', border='B', ln=0, align='C')
        pdf.cell(0, 10, 'Principal Ramakrishna', border='B', ln=1, align='C')
        pdf.set_font('Helvetica', '', 10)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(90, 10, f'Principal, {self.config.ORGANIZATION}', ln=0, align='C')
        pdf.cell(0, 10, f'Principal, {self.config.ORGANIZATION}W', ln=1, align='C')

        # Ensure certificate folder exists
        os.makedirs(self.config.CERTIFICATE_FOLDER, exist_ok=True)
        
        cert_path = os.path.join(self.config.CERTIFICATE_FOLDER, f"{roll_number}.pdf")
        pdf.output(cert_path)
        
        return cert_path
