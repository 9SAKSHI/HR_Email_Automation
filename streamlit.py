import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os

class EmailAutomationApp:
    def __init__(self):
        """
        Initialize the Streamlit Email Automation Application
        """
        st.set_page_config(
            page_title="Job Offer Email Sender", 
            page_icon="✉️", 
            layout="wide"
        )
    
    def load_templates_from_excel(self, excel_file):
        """
        Load email templates from a separate Excel sheet
        
        Expected Excel columns:
        - Role
        - Email Template
        """
        try:
            templates_df = pd.read_excel(excel_file, sheet_name='Templates')
            templates_dict = dict(zip(templates_df['Role'], templates_df['Email Template']))
            return templates_dict
        except Exception as e:
            st.error(f"Error loading templates: {e}")
            return {}
    
    def configure_email_settings(self):
        """
        Simple email configuration interface
        """
        st.header("📧 Email Configuration")
        
        col1, col2 = st.columns(2)
        
        with col1:
            smtp_server = st.text_input("SMTP Server", placeholder="smtp.gmail.com")
            sender_email = st.text_input("Sender Email", placeholder="your.email@company.com")
        
        with col2:
            smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
            sender_password = st.text_input("Email Password", type="password")
        
        return {
            'smtp_server': smtp_server,
            'smtp_port': smtp_port,
            'sender_email': sender_email,
            'sender_password': sender_password
        }
    
    def upload_candidate_data(self):
        """
        Upload and process candidate data
        """
        st.header("📊 Candidate Data")
        
        # File uploaders
        candidate_file = st.file_uploader(
            "Upload Candidate Data Excel", 
            type=['xlsx', 'xls']
        )
        templates_file = st.file_uploader(
            "Upload Email Templates Excel", 
            type=['xlsx', 'xls']
        )
        
        if candidate_file and templates_file:
            # Load candidate data
            candidates_df = pd.read_excel(candidate_file)
            
            # Load templates
            templates = self.load_templates_from_excel(templates_file)
            st.session_state.email_templates = templates
            
            # Display candidate data
            st.subheader("Candidate Information")
            st.dataframe(candidates_df)
            
            # Filter candidates
            status_filter = st.multiselect(
                "Select Status to Send Emails", 
                options=candidates_df['status'].unique(),
                default=['offered']
            )
            
            # Filter dataframe
            filtered_df = candidates_df[candidates_df['status'].isin(status_filter)]
            st.write(f"Candidates Selected: {len(filtered_df)}")
            
            return filtered_df
        
        return None
    
    def send_emails(self, candidates, email_config):
        """
        Automated email sending process
        """
        st.header("✉️ Email Sending Process")
        
        # Create tracking dataframe
        tracking_df = pd.DataFrame(columns=[
            'Candidate Name', 'Email', 'Role', 
            'Send Date', 'Send Time', 'Status', 'Remarks'
        ])
        
        # Progress tracking
        progress_bar = st.progress(0)
        status_placeholder = st.empty()
        
        # SMTP Connection
        try:
            server = smtplib.SMTP(
                email_config['smtp_server'], 
                email_config['smtp_port']
            )
            server.starttls()
            server.login(
                email_config['sender_email'], 
                email_config['sender_password']
            )
        except Exception as e:
            st.error(f"SMTP Connection Error: {e}")
            return
        
        # Email sending loop
        for index, candidate in candidates.iterrows():
            try:
                # Prepare email
                msg = MIMEMultipart()
                msg['From'] = email_config['sender_email']
                msg['To'] = candidate['email']
                msg['Subject'] = f"Offer Letter - {candidate['role'].capitalize()} Position"
                
                # Select template based on role
                template = st.session_state.email_templates.get(
                    candidate['role'], 
                    st.session_state.email_templates.get('default', '')
                )
                
                # Personalize template
                personalized_body = template.format(**candidate)
                msg.attach(MIMEText(personalized_body, 'plain'))
                
                # Send email
                server.send_message(msg)
                
                # Log successful send
                tracking_df = tracking_df.append({
                    'Candidate Name': candidate['name'],
                    'Email': candidate['email'],
                    'Role': candidate['role'],
                    'Send Date': datetime.now().strftime('%Y-%m-%d'),
                    'Send Time': datetime.now().strftime('%H:%M:%S'),
                    'Status': 'Sent',
                    'Remarks': 'Email sent successfully'
                }, ignore_index=True)
                
                # Update progress
                progress_bar.progress((index + 1) / len(candidates))
                status_placeholder.success(f"Sent email to {candidate['name']}")
            
            except Exception as e:
                # Log failed send
                tracking_df = tracking_df.append({
                    'Candidate Name': candidate['name'],
                    'Email': candidate['email'],
                    'Role': candidate['role'],
                    'Send Date': datetime.now().strftime('%Y-%m-%d'),
                    'Send Time': datetime.now().strftime('%H:%M:%S'),
                    'Status': 'Failed',
                    'Remarks': str(e)
                }, ignore_index=True)
                
                status_placeholder.error(f"Failed to send email to {candidate['name']}: {e}")
        
        # Close SMTP connection
        server.quit()
        
        # Save tracking file
        tracking_filename = f"email_tracking_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        tracking_df.to_excel(tracking_filename, index=False)
        st.success(f"Email tracking saved to {tracking_filename}")
        
        # Display tracking results
        st.dataframe(tracking_df)
    
    def run(self):
        """
        Main application runner
        """
        st.title("🚀 Job Offer Email Automation")
        
        # Email Configuration
        email_config = self.configure_email_settings()
        
        # Candidate Data and Templates
        candidates = self.upload_candidate_data()
        
        # Send Emails Button
        if st.button("🚀 Send Offer Emails", key="send_emails_btn"):
            if candidates is not None:
                self.send_emails(candidates, email_config)
            else:
                st.warning("Please upload candidate data and email templates first!")

# Run the Streamlit App
if __name__ == "__main__":
    app = EmailAutomationApp()
    app.run()