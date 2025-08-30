import pandas as pd
import os
import sys
from certificate_generator import CertificateGenerator
from email_sender import EmailSender
from config import Config
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('certificate_sender.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class CertificateProcessor:
    def __init__(self):
        self.config = Config()
        self.certificate_gen = CertificateGenerator()
        self.email_sender = EmailSender()
        self.results = []
    
    def load_participants(self):
        """Load participant data from Excel file"""
        try:
            if not os.path.exists(self.config.EXCEL_FILE):
                raise FileNotFoundError(f"Excel file not found: {self.config.EXCEL_FILE}")
            
            data = pd.read_excel(self.config.EXCEL_FILE, header=1)
            logger.info(f"Loaded {len(data)} participants from Excel file")
            return data
        except Exception as e:
            logger.error(f"Failed to load Excel file: {str(e)}")
            raise
    
    def process_participant(self, participant_data):
        """Process a single participant"""
        try:
            name = str(participant_data.get('NAMES', '')).strip()
            roll = str(participant_data.get('ROLL NUMBERS', '')).strip()
            email = str(participant_data.get('EMAILS', '')).strip()
            
            if not all([name, roll, email]):
                return {
                    'name': name or 'Unknown',
                    'roll': roll or 'Unknown',
                    'email': email or 'Unknown',
                    'status': 'failed',
                    'reason': 'Missing required data'
                }
            
            # Generate certificate
            cert_path = self.certificate_gen.generate_certificate(
                name, roll, self.config.EVENT_NAME, self.config.EVENT_DATE
            )
            
            # Send email
            success, message = self.email_sender.send_certificate(
                email, name, cert_path, self.config.EVENT_NAME
            )
            
            return {
                'name': name,
                'roll': roll,
                'email': email,
                'status': 'success' if success else 'failed',
                'reason': message
            }
            
        except Exception as e:
            logger.error(f"Error processing participant {participant_data}: {str(e)}")
            return {
                'name': str(participant_data.get('NAMES', 'Unknown')),
                'roll': str(participant_data.get('ROLL NUMBERS', 'Unknown')),
                'email': str(participant_data.get('EMAILS', 'Unknown')),
                'status': 'failed',
                'reason': str(e)
            }
    
    def process_all_participants(self):
        """Process all participants"""
        try:
            # Test email connection first
            if not self.email_sender.test_connection():
                logger.error("Email connection test failed. Please check your credentials.")
                return
            
            participants = self.load_participants()
            total = len(participants)
            
            logger.info(f"Starting to process {total} participants...")
            
            for index, participant in participants.iterrows():
                logger.info(f"Processing {index + 1}/{total}: {participant.get('NAMES', 'Unknown')}")
                result = self.process_participant(participant)
                self.results.append(result)
            
            self.generate_report()
            
        except Exception as e:
            logger.error(f"Error processing participants: {str(e)}")
            raise
    
    def generate_report(self):
        """Generate processing report"""
        total = len(self.results)
        successful = len([r for r in self.results if r['status'] == 'success'])
        failed = total - successful
        
        print("\n" + "="*50)
        print("CERTIFICATE SENDING REPORT")
        print("="*50)
        print(f"Total Participants: {total}")
        print(f"Successfully Sent: {successful}")
        print(f"Failed: {failed}")
        print("="*50)
        
        if failed > 0:
            print("\nFailed Emails:")
            for result in self.results:
                if result['status'] == 'failed':
                    print(f"- {result['name']} ({result['email']}): {result['reason']}")
    
    def get_failed_emails(self):
        """Get list of failed emails for retry"""
        return [r for r in self.results if r['status'] == 'failed']

def main():
    """Main function to run the certificate processing"""
    try:
        processor = CertificateProcessor()
        processor.process_all_participants()
    except KeyboardInterrupt:
        logger.info("Process interrupted by user")
    except Exception as e:
        logger.error(f"Process failed: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
