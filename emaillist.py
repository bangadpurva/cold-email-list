#%%
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from dotenv import load_dotenv

#%%
def read_excel(file_path):
    # Read the Excel file
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

#%%
def send_email(recipient_email, recipient_name, staffing_org, attachment_path):
    # Load email credentials from environment variables (.env)
    sender_email = os.getenv('EMAIL_USER')
    sender_password = os.getenv('EMAIL_PASS')

    if not sender_email or not sender_password:
        print("Email credentials not found in environment variables.")
        return

    # Email content
    subject = f"Exploring [your desired position] Opportunities at {staffing_org}"
    body = f"""
    <p>Hey {recipient_name},</p>

    <p>[Add your skills and contribution]</p>

    <p>Best regards,</p>
    <p>[Your Name]</p>
    <p><a href="https://linkedin.com/in/your_profile">LinkedIn</a></p>
    """

    try:
        # Set up the MIME
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'html'))

        # Attach the file
        if attachment_path:
            try:
                attachment = open(attachment_path, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
                message.attach(part)
            except Exception as e:
                print(f"Error attaching file: {e}")

        # Log in to the server
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        # server.starttls()
        server.login(sender_email, sender_password)

        # Send the email
        server.sendmail(sender_email, recipient_email, message.as_string())
        server.quit()

        print(f"Email sent to {recipient_name} for {staffing_org} at ({recipient_email})")
    except smtplib.SMTPException as e:
        print(f"Error sending email: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

#%%
def main():
    
    # Load environment variables from .env file
    load_dotenv()

    file_path = 'Recs.xlsx'  # Update with your file path with xlsz format list of recruiters
    attachment_path = 'Resume.pdf'  # Update with your resume attachment path
    df = read_excel(file_path)

    if df is not None:
        # Iterate through the DataFrame and send emails
        for index, row in df.iterrows():
            recipient_name = row['Name']
            staffing_org = row['Staffing Organisation']
            recipient_email = row['Email']
            send_email(recipient_email, recipient_name, staffing_org, attachment_path)

#%%
if __name__ == "__main__":
    main()
