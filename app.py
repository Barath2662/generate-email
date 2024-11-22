from flask import Flask, request, render_template, send_file
import os
import pandas as pd
from pptx import Presentation
import zipfile
from flask_mail import Mail, Message

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

app.config['MAIL_SERVER'] = 'smtp.gmail.com' 
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'barathvikraman.projects@gmail.com'  
app.config['MAIL_PASSWORD'] = 'barath2606'  
mail = Mail(app)

def modify_certificate(template_path, name, output_path):
    prs = Presentation(template_path)
    

    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                shape.text = shape.text.replace('{{Name}}', name)
    
  
    prs.save(output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_certificates():
    if 'template' not in request.files or 'data' not in request.files:
        return render_template('index.html', message='No file part')

    template_file = request.files['template']
    data_file = request.files['data']

    if template_file.filename == '' or data_file.filename == '':
        return render_template('index.html', message='No selected file')


    email_subject = request.form.get('subject')
    email_body = request.form.get('body')


    template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file.filename)
    data_path = os.path.join(app.config['UPLOAD_FOLDER'], data_file.filename)
    
    template_file.save(template_path)
    data_file.save(data_path)


    df = pd.read_excel(data_path)


    zip_filename = 'certificates.zip'
    zip_filepath = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)

    with zipfile.ZipFile(zip_filepath, 'w') as zipf:
        for index, row in df.iterrows():
            name = row['Name']  
            
            output_filename = f'certificate_{index + 1}.pptx'
            output_filepath = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            

            modify_certificate(template_path, name, output_filepath)
            

            zipf.write(output_filepath, arcname=output_filename)

            recipient_email = row['Email']  # Adjust based on your Excel column name
            send_email(recipient_email, email_subject, email_body, output_filepath)

    os.remove(template_path)
    os.remove(data_path)

    return send_file(zip_filepath, as_attachment=True)

def send_email(recipient_email, subject, body, attachment_path):
    msg = Message(
        subject=subject,
        sender=app.config['MAIL_USERNAME'],
        recipients=[recipient_email],  # Send to recipient's email address from Excel
        body=body.format(name=recipient_email.split('@')[0])  # You can customize this if needed.
    )
    
    with app.open_resource(attachment_path) as fp:
        msg.attach(os.path.basename(attachment_path), 'application/vnd.openxmlformats-officedocument.presentationml.presentation', fp.read())

    mail.send(msg)

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
        
    app.run(debug=True)