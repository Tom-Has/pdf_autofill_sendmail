import PyPDF2

def get_pdf_form_fields(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        fields = pdf_reader.get_fields()
        return fields

def print_form_fields(fields):
    for field_name, field_info in fields.items():
        print(f"Field Name: {field_name}")
        for key, value in field_info.items():
            print(f"  {key}: {value}")
        print()

pdf_path = "path_to_file/filename.pdf"
fields = get_pdf_form_fields(pdf_path)
print_form_fields(fields)

