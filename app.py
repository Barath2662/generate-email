import streamlit as st
import pandas as pd
from pptx import Presentation
import zipfile
import os
from flask_mail import Mail, Message

# Configure Flask-Mail settings (modify with your credentials)
MAIL_SERVER = 'smtp.gmail.com'
MAIL_PORT = 587
MAIL_USE_TLS = True
MAIL_USERNAME = 'barathvikraman.projects@gmail.com'  # Your email address
MAIL_PASSWORD = 'barath2606'          # Your email password or app password

# Initialize Flask-Mail (not used directly in Streamlit but setup for sending emails)
mail = Mail()

def modify_certificate(template_path, name, output_path):
    prs = Presentation(template_path)
    
    # Replace {{Name}} placeholder in each slide of the presentation
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                shape.text = shape.text.replace('{{Name}}', name)
    
    # Save the modified PPTX file
    prs.save(output_path)

def send_email(recipient_email, subject, body, attachment_path):
    msg = Message(
        subject=subject,
        sender=MAIL_USERNAME,
        recipients=[recipient_email],
        body=body.format(name=recipient_email.split('@')[0])
    )
    
    with open(attachment_path, "rb") as fp:
        msg.attach(os.path.basename(attachment_path), 'application/vnd.openxmlformats-officedocument.presentationml.presentation', fp.read())

    mail.send(msg)

def main():
    st.title("Certificate Generator")

    # File upload section
    template_file = st.file_uploader("Upload PPTX Template", type=["pptx"])
    data_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    email_subject = st.text_input("Email Subject")
    email_body = st.text_area("Email Body")

    if st.button("Generate Certificates"):
        if template_file and data_file:
            # Save uploaded files temporarily
            template_path = os.path.join('uploads', template_file.name)
            data_path = os.path.join('uploads', data_file.name)

            with open(template_path, "wb") as f:
                f.write(template_file.getbuffer())
            with open(data_path, "wb") as f:
                f.write(data_file.getbuffer())

            # Read Excel data
            df = pd.read_excel(data_path)

            # Create a zip file to store all generated PPTX certificates
            zip_filename = 'certificates.zip'
            zip_filepath = os.path.join('uploads', zip_filename)

            with zipfile.ZipFile(zip_filepath, 'w') as zipf:
                for index, row in df.iterrows():
                    name = row['Name']  # Adjust based on your Excel column name
                    
                    output_filename = f'certificate_{index + 1}.pptx'
                    output_filepath = os.path.join('uploads', output_filename)
                    
                    # Modify the PPTX template with the name
                    modify_certificate(template_path, name, output_filepath)
                    
                    # Add the modified PPTX to the zip file
                    zipf.write(output_filepath, arcname=output_filename)

                    # Send email with attachment (assuming there's an Email column in Excel)
                    recipient_email = row['Email']  # Adjust based on your Excel column name
                    send_email(recipient_email, email_subject, email_body, output_filepath)

            st.success(f"Certificates generated and sent! Download ZIP: {zip_filename}")
            st.download_button("Download ZIP", zip_filepath)

if __name__ == "__main__":
    main()
