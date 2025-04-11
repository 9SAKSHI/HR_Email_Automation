import pandas as pd
import time
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ExcelChangeHandler(FileSystemEventHandler):
    def __init__(self, email_sender):
        self.email_sender = email_sender
        self.last_modified_candidates = {}
    
    def on_modified(self, event):
        if not event.is_directory and event.src_path.endswith(('.xlsx', '.xls')):
            try:
                # Wait a moment to ensure file is fully written
                time.sleep(1)
                
                # Load current candidates data
                current_candidates = pd.read_excel(event.src_path)
                
                # Compare with last known state
                for index, candidate in current_candidates.iterrows():
                    # Create a unique identifier for the candidate
                    candidate_key = f"{candidate['name']}_{candidate['email']}"
                    
                    # Check if status has changed to 'offered'
                    if (candidate['status'] == 'offered' and 
                        (candidate_key not in self.last_modified_candidates or 
                         self.last_modified_candidates[candidate_key] != 'offered')):
                        
                        # Send offer email
                        self.email_sender.send_offer_email(candidate)
                        
                        # Update last known state
                        self.last_modified_candidates[candidate_key] = 'offered'
                
                logging.info(f"Processed changes in {event.src_path}")
            
            except Exception as e:
                logging.error(f"Error processing file change: {e}")

class EmailAutomationSystem:
    def __init__(self, config_path):
        # Setup logging
        logging.basicConfig(
            filename='email_automation.log', 
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s: %(message)s'
        )
        
        # Load configuration
        self.load_configuration(config_path)
        
        # Load email templates
        self.email_templates = self.load_email_templates()
    
    def load_configuration(self, config_path):
        """
        Load configuration from Excel
        """
        try:
            config_df = pd.read_excel(config_path, sheet_name='EmailConfig')
            
            # Email Server Configuration
            self.smtp_server = config_df.loc[config_df['Key'] == 'SMTP_SERVER', 'Value'].values[0]
            self.smtp_port = int(config_df.loc[config_df['Key'] == 'SMTP_PORT', 'Value'].values[0])
            self.sender_email = config_df.loc[config_df['Key'] == 'SENDER_EMAIL', 'Value'].values[0]
            self.sender_password = config_df.loc[config_df['Key'] == 'SENDER_PASSWORD', 'Value'].values[0]
            
            # Paths Configuration
            self.candidates_file = config_df.loc[config_df['Key'] == 'CANDIDATE_FILE', 'Value'].values[0]
            self.template_file = config_df.loc[config_df['Key'] == 'TEMPLATE_FILE', 'Value'].values[0]
            self.tracking_folder = config_df.loc[config_df['Key'] == 'TRACKING_FOLDER', 'Value'].values[0]
            
            logging.info("Configuration loaded successfully")
        
        except Exception as e:
            logging.error(f"Configuration load error: {e}")
            raise
    
    def load_email_templates(self):
        """
        Load email templates from Excel
        """
        try:
            templates_df = pd.read_excel(self.template_file, sheet_name='Templates')
            return dict(zip(templates_df['Role'], templates_df['Template']))
        except Exception as e:
            logging.error(f"Template loading error: {e}")
            return {}
    
    def personalize_email_template(self, template, candidate):
        """
        Personalize email template with candidate details
        """
        try:
            # Add custom personalizations
            personalization_map = {
                '{name}': candidate['name'],
                '{email}': candidate['email'],
                '{role}': candidate['role'],
                '{department}': candidate.get('department', 'N/A'),
                '{start_date}': candidate.get('start_date', 'TBD'),
                '{location}': candidate.get('location', 'Company Location')
            }
            
            # Replace placeholders
            for placeholder, value in personalization_map.items():
                template = template.replace(placeholder, str(value))
            
            return template
        
        except Exception as e:
            logging.error(f"Template personalization error: {e}")
            return template
    
    def send_offer_email(self, candidate):
        """
        Send offer email to a candidate
        """
        try:
            # Create SMTP connection
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                
                # Prepare email
                msg = MIMEMultipart()
                msg['From'] = self.sender_email
                msg['To'] = candidate['email']
                msg['Subject'] = f"Offer Letter - {candidate['role'].capitalize()} Position"
                
                # Select and personalize template
                template = self.email_templates.get(
                    candidate['role'], 
                    self.email_templates.get('default', 'Congratulations on your offer!')
                )
                personalized_body = self.personalize_email_template(template, candidate)
                
                # Attach email body
                msg.attach(MIMEText(personalized_body, 'plain'))
                
                # Send email
                server.send_message(msg)
                
                # Log successful send
                logging.info(f"Offer email sent to {candidate['name']} for {candidate['role']} role")
                
                # Create tracking record
                self.create_tracking_record(candidate)
        
        except Exception as e:
            logging.error(f"Email send error for {candidate['name']}: {e}")
    
    def create_tracking_record(self, candidate):
        """
        Create a tracking record for sent emails
        """
        try:
            # Prepare tracking dataframe
            tracking_df = pd.DataFrame([{
                'Candidate Name': candidate['name'],
                'Email': candidate['email'],
                'Role': candidate['role'],
                'Send Date': pd.Timestamp.now(),
                'Status': 'Sent',
                'Remarks': 'Offer email sent successfully'
            }])
            
            # Save to tracking file
            tracking_filename = os.path.join(
                self.tracking_folder, 
                f"email_tracking_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            tracking_df.to_excel(tracking_filename, index=False)
        
        except Exception as e:
            logging.error(f"Tracking record creation error: {e}")

def main():
    # Initialize email automation system
    email_sender = EmailAutomationSystem('config.xlsx')
    
    # Create file change handler
    event_handler = ExcelChangeHandler(email_sender)
    
    # Create observer
    observer = Observer()
    observer.schedule(
        event_handler, 
        path=os.path.dirname(email_sender.candidates_file), 
        recursive=False
    )
    
    # Start monitoring
    observer.start()
    
    try:
        logging.info("Excel change monitoring started")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    
    observer.join()

if __name__ == "__main__":
    main()