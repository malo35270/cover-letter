import argparse
from docx import Document
from datetime import date
from docx2pdf import convert
from python_docx_replace import docx_replace
import os

def generate_cl(args):
    document = Document(f'source/cl-{args.lan}.docx')
    my_dict ={
        'Position' : args.position,
        "Recipient's Name" : args.recipient,
        "Company's Name" : args.company_name,
        "Date" : date.today().isoformat(),
        "Company's Address" : args.company_address if args.company_address else '',
        "City, State, ZIP Code": args.company_city if args.company_city else '',
    }
    docx_replace(document, **my_dict)

    path = os.path.join('result', args.company_name)
        
        # Create the folder in the specified directory
    print(path)
    os.makedirs(path, exist_ok=True)
    document.save(f'{path}/cover-letter-Malo-Chauvel-{args.position}.docx')
    convert(f'{path}/cover-letter-Malo-Chauvel-{args.position}.docx', f'{path}/cover-letter-Malo-Chauvel-{args.position}.pdf')


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-lan", help="langue of the cover letter")
    parser.add_argument("-p","--position", help="position in the offer")
    parser.add_argument("-r","--recipient", help="recipient : who should address the letter")
    parser.add_argument("-cn","--company_name", help="company_name")
    parser.add_argument("--company_address", help="company_address")
    parser.add_argument("--company_city", help="City, State, ZIP Code")
    args = parser.parse_args()
    generate_cl(args)