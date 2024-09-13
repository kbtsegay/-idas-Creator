import os
import base64
import yagmail
from argparse import ArgumentParser
from src.kidase_creator import KidaseCreator

def decode_credentials(encoded):
    # Decode the base64 string
    decoded = base64.b64decode(encoded)
    # Write the decoded content to credentials.json
    with open('credentials.json', 'wb') as file:
        file.write(decoded)

def send_email(file_path, recipient_email):
    # Initialize yagmail with OAuth2 credentials
    yag = yagmail.SMTP("kidasecreator.noreply@gmail.com", oauth2_file="./credentials.json")
    
    # Create the email content
    subject = "Your Presentation"
    body = "Please find the attached presentation."
    
    # Send the email with the attachment
    yag.send(
        to=recipient_email,
        subject=subject,
        contents=body,
        attachments=file_path
    )

# Example usage
if __name__ == '__main__':
    parser = ArgumentParser(description='Decode the base64 encoded credentials')
    parser.add_argument('encoded', type=str, help='The base64 encoded credentials')
    args = parser.parse_args()

    print('Decoding credentials...')
    decode_credentials(args.encoded)

    kidase_creator = KidaseCreator('./data', ['ግእዝ', 'ትግርኛ', 'english'])
    prs = kidase_creator.create_presentation()
    prs.save('test.pptx')

    send_email('test.pptx', 'kaleb.tsegay@gmail.com')
