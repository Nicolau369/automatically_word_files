import pandas as pd
from docx import Document

def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
                for run in paragraph.runs:
                    run.text = run.text.replace(key, value)

    doc.save(output_path)

def generate_invitations_from_csv(csv_path, template_path):
    df = pd.read_csv('csv.path')
    for idx, row in df.iterrows():
        data = {
            '[Salutation]': row['Salutation'],
            '[First Name]': row['First_Name'],
            '[Last Name]': row['Last_Name'],
            '[Let Connect]': row['Let_Connect'],
            '[Company Name]': row['Company']
        }
        output_path = f'invitation_{idx + 1}.docx'
        fill_invitation(template_path, output_path, data)


if __name__ == '__main__':
   csv_path = 'contacts.csv',
   template_path = '1rma0lh0.docx',
   generate_invitations_from_csv(csv_path, template_path)