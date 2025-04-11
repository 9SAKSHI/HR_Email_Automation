import streamlit as st
import pandas as pd
import os
import datetime
import win32com.client as win32
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import io
import time
from pathlib import Path

# Set page configuration
st.set_page_config(page_title="Email Automation System", layout="wide")

# Initialize session state variables if they don't exist
if 'monitoring_active' not in st.session_state:
    st.session_state.monitoring_active = False
if 'last_check_time' not in st.session_state:
    st.session_state.last_check_time = None
if 'emails_sent' not in st.session_state:
    st.session_state.emails_sent = []
if 'template_paths' not in st.session_state:
    st.session_state.template_paths = {}
if 'tracking_df' not in st.session_state:
    st.session_state.tracking_df = pd.DataFrame(columns=['Candidate Name', 'Email', 'Type', 'Location', 'Sent Time', 'Status'])
if 'df' not in st.session_state:
    st.session_state.df = None

# Define candidate types and their corresponding template URLs
TEMPLATE_URLS = {
    "Intern": r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Intern\Pune\Gear-up for your exciting journey with Dassault Systemes! (Intern).msg",
    "Direct Contractor":r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Direct Contractor Fresher\Pune\Gear-up for your exciting journey with Dassault Systemes! (Fresher).msg",
    "Direct Contractor Lateral": r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Direct Contractor Lateral\Banglore\Gear-up for your exciting journey with Dassault Systemes! (Lateral).msg",
    "Apprentice":r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Apprentice\Pune\Gear-up for your exciting journey with Dassault Systemes! (Apprentice).msg",
    "Regular Fresher": r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Regular Fresher\Pune\Gear-up for your exciting journey with Dassault Systemes! (Fresher).msg",
    "Regular Lateral": r"C:\Users\SNR23\Desktop\HR Automation\Email_Automation\Gearup Email Automation\Gearup Email Automation\DSGS\Regular Lateral\Pune\Gear-up for your exciting journey with Dassault Systemes! (Lateral).msg"
}

def send_email_from_template(row, template_path):
    """Send email using Outlook template"""
    try:
        # Extract candidate information
        candidate_name = row['Name']
        email = row['Candidate Email Id']
        location =row['Location']
        candidate_type = row['Emp Type']
        
        # Get joining date (2 weeks from now by default)
        joining_date = candidate_row['DOJ']
        
        # Create Outlook application object
        outlook = win32.Dispatch('Outlook.Application')
        
        # Open the template
        mail = outlook.CreateItemFromTemplate(template_path)
        
        # Set email properties
        mail.To = email
        
        # Replace placeholders in the email body
        mail_body = mail.Body
        mail_body = mail_body.replace("candidate name", candidate_name)
        mail_body = mail_body.replace('"date of joining"', joining_date)
        mail_body = mail_body.replace('"location"', location)
        mail_body = mail_body.replace("company name", "Dassault Systemes")
        
        # Update the email body
        mail.Body = mail_body
        
        # Send the email
        mail.Send()
        
        # Log the successful email
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_row = pd.DataFrame({
            'Candidate Name': [candidate_name],
            'Email': [email],
            'Type': [candidate_type],
            'Location': [location],
            'Sent Time': [current_time],
            'Status': ['Sent']
        })
        
        st.session_state.tracking_df = pd.concat([st.session_state.tracking_df, new_row], ignore_index=True)
        st.session_state.emails_sent.append(f"{candidate_name} - {email} - {current_time}")
        
        return True, f"Email sent successfully to {candidate_name}"
        
    except Exception as e:
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        error_msg = f"Error sending email to {candidate_row['Name']}: {str(e)}"
        
        # Log the failed email
        new_row = pd.DataFrame({
            'Candidate Name': [candidate_row['Name']],
            'Email': [candidate_row['Candidate Email Id']],
            'Type': [candidate_row['Emp Type']],
            'Location': [candidate_row['Location']],
            'Sent Time': [current_time],
            'Status': [f"Failed: {str(e)}"]
        })
        
        st.session_state.tracking_df = pd.concat([st.session_state.tracking_df, new_row], ignore_index=True)
        
        return False, error_msg

def check_for_new_offers(df, template_paths):
    """Check for new candidates with 'Offered' status"""
    if df is None or df.empty:
        return "No data available."
    
    # Find rows with status "Offered" 
    new_offers = df[df['Status'] == 'Offered'].copy()
    
    if len(new_offers) == 0:
        return "No new offers to process."
        
    # Process each new offer
    results = []
    for idx, row in new_offers.iterrows():
        candidate_name = row['Name']
        candidate_type = row['Emp Type']
        
        # Check if we have a template for this candidate type
        if candidate_type not in template_paths or not template_paths[candidate_type]:
            results.append(f"No template available for {candidate_name} ({candidate_type})")
            continue
            
        # Send the email
        success, message = send_email_from_template(row, template_paths[candidate_type])
        results.append(message)
        
        # Update the dataframe to mark as processed
        if success:
            df.at[idx, 'Email Sent'] = 'Yes'
            df.at[idx, 'Email Sent Date'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    return "\n".join(results)

# Watchdog Handler Class
class FileChangeHandler(FileSystemEventHandler):
    def __init__(self, df, template_paths):
        self.df = df
        self.template_paths = template_paths
        
    def on_modified(self, event):
        """Called when the file is modified"""
        if event.src_path.endswith('.xlsx'):
            print("File modified, reloading data...")
            self.df = pd.read_excel(event.src_path)  # Reload the Excel file
            # Look for new 'Offered' candidates
            result = check_for_new_offers(self.df, self.template_paths)
            if result:
                st.text(result)

def monitor_excel_file(file_path):
    """Start the watchdog observer to monitor changes in the file"""
    event_handler = FileChangeHandler(st.session_state.df, st.session_state.template_paths)
    observer = Observer()
    observer.schedule(event_handler, path=os.path.dirname(file_path), recursive=False)
    observer.start()
    return observer

# Main app UI
st.title("Email Automation System")

# Sidebar for configuration
st.sidebar.header("Configuration")

# File upload section
st.sidebar.subheader("Upload Candidate Data")
uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.session_state.df = df
        
        # Ensure required columns exist
        if 'Email Sent' not in df.columns:
            df['Email Sent'] = None
        if 'Email Sent Date' not in df.columns:
            df['Email Sent Date'] = None
            
        st.sidebar.success("File uploaded successfully!")
        
        # Start file monitoring in a background thread
        st.session_state.observer = monitor_excel_file(uploaded_file.name)
    except Exception as e:
        st.sidebar.error(f"Error: {str(e)}")

# Set the template paths from the predefined URLs
st.session_state.template_paths = TEMPLATE_URLS

# Main content area - Tabs
tab1, tab2, tab3, tab4 = st.tabs(["Candidate Data", "Send Emails", "Email History", "Instructions"])

# Tab 1: Candidate Data
with tab1:
    st.header("Candidate Data")
    if 'df' in st.session_state and not st.session_state.df.empty:
        st.dataframe(st.session_state.df, use_container_width=True)
        
        # Add a download button for the updated Excel
        if st.button("Download Updated Excel"):
            excel_data = io.BytesIO()
            with pd.ExcelWriter(excel_data, engine='openpyxl') as writer:
                st.session_state.df.to_excel(writer, index=False)
            excel_data.seek(0)
            
            st.download_button(
                label="Download Excel File",
                data=excel_data,
                file_name="updated_candidates.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Please upload an Excel file with candidate data")

# Tab 2: Send Emails
with tab2:
    st.header("Send Emails")
    
    if 'df' in st.session_state and not st.session_state.df.empty:
        # Filter for candidates with "Offered" status
        offered_candidates = st.session_state.df[st.session_state.df['Status'] == 'Offered']
        
        if not offered_candidates.empty:
            st.subheader("Candidates with 'Offered' Status")
            st.dataframe(offered_candidates[['Name', 'Emp Type', 'Location', 'Candidate Email Id']], use_container_width=True)
            
            # Manual send options
            st.subheader("Send Emails")
            
            if st.session_state.template_paths:
                if st.button("Send Emails to All Offered Candidates"):
                    results = check_for_new_offers(st.session_state.df, st.session_state.template_paths)
                    st.text(results)
