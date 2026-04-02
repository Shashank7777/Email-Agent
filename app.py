import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from datetime import datetime
import os
import markdown
import time
import mimetypes
from dotenv import load_dotenv

# NEW: Import your config file
try:
    import config
except ImportError:
    st.error("Could not find config.py. Please create it in the same directory as this script.")
    st.stop()

# --- Load Environment Variables ---
load_dotenv()

# --- Configuration & Setup ---
st.set_page_config(page_title="Professor Email Automator", page_icon="✉️", layout="wide")
LOG_FILE = "sent_emails_log.xlsx"
ATTACHMENTS_DIR = "attachments"

# Create attachments directory if it doesn't exist
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

# --- Helper Functions ---
def verify_login(email, password):
    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(email, password)
        server.quit()
        return True, "Connection successful! ✅"
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check your email and App Password. ❌"
    except Exception as e:
        return False, f"Connection error: {str(e)} ❌"

def send_email(sender_email, app_password, recipient_email, subject, md_body, html_body, attachment_paths=None):
    if attachment_paths is None:
        attachment_paths =[]
        
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender_email
    msg['To'] = recipient_email
    
    # Set the plain text version first (fallback)
    msg.set_content(md_body)
    # Add the HTML version (rich text)
    msg.add_alternative(html_body, subtype='html')

    # Add attachments
    for file_path in attachment_paths:
        try:
            ctype, encoding = mimetypes.guess_type(file_path)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            
            maintype, subtype = ctype.split('/', 1)
            
            with open(file_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(file_path)
                
            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)
        except Exception as e:
            return False, f"Failed to attach {file_path}: {str(e)}"

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(sender_email, app_password)
        server.send_message(msg)
        server.quit()
        return True, ""
    except Exception as e:
        return False, str(e)

def save_to_log(log_data):
    new_df = pd.DataFrame(log_data)
    if os.path.exists(LOG_FILE):
        existing_df = pd.read_excel(LOG_FILE)
        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        updated_df = new_df
    updated_df.to_excel(LOG_FILE, index=False)

# --- Sidebar for Email Credentials ---
st.sidebar.header("Email Credentials")
st.sidebar.info("Credentials auto-populated from .env file.")

env_email = os.getenv("SENDER_EMAIL", "")
env_password = os.getenv("APP_PASSWORD", "")

sender_email = st.sidebar.text_input("Your Email Address", value=env_email)
app_password = st.sidebar.text_input("App Password", type="password", value=env_password)

if st.sidebar.button("Test Connection"):
    if sender_email and app_password:
        with st.sidebar.spinner("Checking connection..."):
            is_valid, msg = verify_login(sender_email, app_password)
            if is_valid:
                st.sidebar.success(msg)
            else:
                st.sidebar.error(msg)
    else:
        st.sidebar.warning("Please enter credentials first.")

# --- App Layout (Tabs) ---
tab1, tab2 = st.tabs(["📧 Send Emails", "📊 Email Tracker"])

# ==========================================
# TAB 1: SEND EMAILS
# ==========================================
with tab1:
    st.title("✉️ Professor Email Automator")
    st.write("Upload your Excel sheet, customize your template, preview it, and automate your outreach safely.")

    # 1. File Uploads
    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx", "xls"])
        st.caption("Required columns: `Name`, `Research`, `Email`")

    with col2:
        md_file = st.file_uploader("Upload Markdown Template (.md)", type=["md", "txt"])
        st.caption("Uploading populates the editor below.")

    # 2. Email Variables (Prepopulated from config.py)
    st.header("Email Details")
    
    email_subject = st.text_input("Email Subject", value=config.DEFAULT_SUBJECT)
    
    col3, col4 = st.columns(2)
    with col3:
        degree_options = ["PhD", "MS"]
        # Find index dynamically based on config
        default_deg_idx = degree_options.index(config.DEFAULT_DEGREE) if config.DEFAULT_DEGREE in degree_options else 0
        degree = st.selectbox("Applying For", degree_options, index=default_deg_idx)
        
        term = st.text_input("Applied Term", value=config.DEFAULT_TERM)
        university = st.text_input("Applying University", value=config.DEFAULT_UNIVERSITY)
    
    with col4:
        sender_name = st.text_input("Your Name (For 'Regards')", value=config.DEFAULT_SENDER_NAME)
        website = st.text_input("Personal Website URL (Optional)", value=config.DEFAULT_WEBSITE)

    # 3. Editor & Live Preview Section
    st.write("---")
    st.subheader("📝 Edit & Preview Template")
    
    default_template = md_file.getvalue().decode("utf-8") if md_file else ""
    
    editor_tab, preview_tab = st.tabs(["🖋️ Markdown Editor", "👀 Live Email Preview"])
    
    with editor_tab:
        edited_template = st.text_area(
            "Modify your Markdown template here:", 
            value=default_template, 
            height=300
        )
        st.caption("Tags: `<<Name>>`, `<<Research>>`, `<<Degree>>`, `<<Term>>`, `<<University>>`, `<<SenderName>>`, `<<Website>>`")

    with preview_tab:
        if edited_template.strip():
            preview_name = "Dr. Alan Turing"
            preview_research = "Artificial Intelligence"
            
            if excel_file is not None:
                try:
                    df_preview = pd.read_excel(excel_file)
                    if not df_preview.empty and 'Name' in df_preview.columns and 'Research' in df_preview.columns:
                        preview_name = str(df_preview.iloc[0]['Name'])
                        preview_research = str(df_preview.iloc[0]['Research'])
                except Exception:
                    pass

            preview_md = edited_template.replace("<<Name>>", preview_name)
            preview_md = preview_md.replace("<<Research>>", preview_research)
            preview_md = preview_md.replace("<<Degree>>", degree)
            preview_md = preview_md.replace("<<Term>>", term)
            preview_md = preview_md.replace("<<SenderName>>", sender_name)
            preview_md = preview_md.replace("<<University>>", university)
            
            if website:
                preview_md = preview_md.replace("<<Website>>", f"[My Personal Website]({website})")
            else:
                preview_md = preview_md.replace("<<Website>>", "")

            preview_subject = email_subject.replace("<<Degree>>", degree).replace("<<Term>>", term).replace("<<Name>>", preview_name)

            st.info(f"**Subject:** {preview_subject}")
            st.markdown(preview_md, unsafe_allow_html=True)
        else:
            st.warning("Template is empty. Start typing in the Markdown Editor tab!")

    st.write("---")
    
    # 4. Attachments Section
    st.subheader("📎 Attachments")
    st.write(f"Drop your CV or transcripts into the `{ATTACHMENTS_DIR}` folder next to this script.")
    
    available_files =[f for f in os.listdir(ATTACHMENTS_DIR) if os.path.isfile(os.path.join(ATTACHMENTS_DIR, f))]
    
    if available_files:
        selected_files = st.multiselect(
            "Select files to attach to ALL emails:", 
            options=available_files, 
            default=available_files
        )
    else:
        st.info(f"No files found in the '{ATTACHMENTS_DIR}' folder.")
        selected_files =[]

    st.write("---")
    
    # 5. Final Sending Actions (Delay prepopulated from config)
    delay = st.slider("Delay between emails (seconds) - Helps avoid Gmail rate limits", 1, 30, config.DEFAULT_DELAY)
    send_button = st.button("Send Emails 🚀", use_container_width=True, type="primary")

    if send_button:
        if not sender_email or not app_password:
            st.error("Please provide your email and app password in the sidebar.")
        elif excel_file is None:
            st.error("Please upload your Excel file containing the professor contacts.")
        elif not edited_template.strip():
            st.error("The email template is empty!")
        else:
            try:
                df = pd.read_excel(excel_file)
                required_cols = ['Name', 'Research', 'Email']
                
                if not all(col in df.columns for col in required_cols):
                    st.error(f"Excel file must contain exact column names: {', '.join(required_cols)}")
                else:
                    success_count = 0
                    error_list = []
                    logs_to_save =[]
                    
                    total_emails = len(df)
                    status_text = st.empty()
                    progress_bar = st.progress(0)
                    
                    attachment_paths =[os.path.join(ATTACHMENTS_DIR, f) for f in selected_files]
                    
                    for index, row in df.iterrows():
                        prof_name = str(row['Name'])
                        prof_research = str(row['Research'])
                        prof_email = str(row['Email']).strip()
                        prof_university = str(row['University']) if 'University' in df.columns else "N/A"
                        
                        status_text.info(f"Processing email {index + 1} of {total_emails} for {prof_name} ({prof_email})...")
                        
                        personalized_md = edited_template.replace("<<Name>>", prof_name)
                        personalized_md = personalized_md.replace("<<Research>>", prof_research)
                        personalized_md = personalized_md.replace("<<Degree>>", degree)
                        personalized_md = personalized_md.replace("<<Term>>", term)
                        personalized_md = personalized_md.replace("<<SenderName>>", sender_name)
                        personalized_md = personalized_md.replace("<<University>>", university)
                        
                        if website:
                            personalized_md = personalized_md.replace("<<Website>>", f"[My Personal Website]({website})")
                        else:
                            personalized_md = personalized_md.replace("<<Website>>", "")

                        html_content = markdown.markdown(personalized_md)
                        personalized_subject = email_subject.replace("<<Degree>>", degree).replace("<<Term>>", term).replace("<<Name>>", prof_name)
                        
                        is_success, error_msg = send_email(
                            sender_email, 
                            app_password, 
                            prof_email, 
                            personalized_subject, 
                            personalized_md, 
                            html_content,
                            attachment_paths
                        )
                        
                        if is_success:
                            success_count += 1
                            logs_to_save.append({
                                "Date": datetime.now().strftime("%Y-%m-%d"),
                                "Time": datetime.now().strftime("%H:%M:%S"),
                                "Professor Name": prof_name,
                                "Email": prof_email,
                                "University Name": prof_university,
                                "Research Interest": prof_research,
                                "Email Text": personalized_md,
                                "Attachments": ", ".join(selected_files) if selected_files else "None"
                            })
                        else:
                            error_list.append(f"Failed to send to {prof_email}: {error_msg}")
                            
                        progress_bar.progress((index + 1) / total_emails)

                        if index < total_emails - 1:
                            status_text.warning(f"Sent to {prof_name}. Pausing for {delay} seconds to prevent rate limits...")
                            time.sleep(delay)
                    
                    status_text.empty()
                    
                    if logs_to_save:
                        save_to_log(logs_to_save)

                    if success_count == total_emails:
                        st.success(f"Successfully sent all {total_emails} formatted emails with attachments! 🎉")
                    else:
                        st.warning(f"Sent {success_count}/{total_emails} emails.")
                        for err in error_list:
                            st.error(err)
                            
            except Exception as e:
                st.error(f"Error processing files: {e}")

# ==========================================
# TAB 2: EMAIL TRACKER
# ==========================================
with tab2:
    st.title("📊 Sent Emails Tracker")
    st.write("View, search, and filter all successfully sent emails.")

    if os.path.exists(LOG_FILE):
        log_df = pd.read_excel(LOG_FILE)
        log_df = log_df.fillna("") 

        st.subheader("🔍 Search & Filter")
        search_query = st.text_input("Global Search (Type anything: text, email, name...)")
        
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1:
            date_filter = st.multiselect("Filter by Date", options=log_df["Date"].unique())
        with col_f2:
            uni_filter = st.multiselect("Filter by University", options=log_df["University Name"].unique())
        with col_f3:
            research_filter = st.multiselect("Filter by Research", options=log_df["Research Interest"].unique())
        with col_f4:
            prof_filter = st.multiselect("Filter by Professor", options=log_df["Professor Name"].unique())

        filtered_df = log_df.copy()

        if search_query:
            mask = filtered_df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)
            filtered_df = filtered_df[mask]

        if date_filter:
            filtered_df = filtered_df[filtered_df["Date"].isin(date_filter)]
        if uni_filter:
            filtered_df = filtered_df[filtered_df["University Name"].isin(uni_filter)]
        if research_filter:
            filtered_df = filtered_df[filtered_df["Research Interest"].isin(research_filter)]
        if prof_filter:
            filtered_df = filtered_df[filtered_df["Professor Name"].isin(prof_filter)]

        st.write(f"**Total Records Found: {len(filtered_df)}**")
        st.dataframe(filtered_df, use_container_width=True, height=500)
        
        csv = filtered_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="⬇️ Download Filtered Logs as CSV",
            data=csv,
            file_name='filtered_email_logs.csv',
            mime='text/csv',
        )

    else:
        st.info("No logs found yet. Send some emails from the other tab to create the tracking file!")