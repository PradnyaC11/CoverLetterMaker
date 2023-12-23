from docx import Document
from datetime import date
from docx2pdf import convert


def main():
    template_file_path = "InternshipTemplate.docx"
    output_file_path = "D:/Resumes/Internship/Pradnya_Chaudhari_Cover_Letter.docx"
    pdf_file_path = "D:/Resumes/Internship/Pradnya_Chaudhari_Cover_Letter.pdf"

    today = date.today()
    formatted_date = today.strftime('%d %B, %Y')

    variables = {
        "${DATE}": formatted_date,
        "${COMPANY}": "Nishtha Incorporated",
        "${ADDRESS}": "Tempe, AZ, United States",
        "${POSITION}": "Cloud Software Development Intern",
        "${SHORTPOS}": "Cloud Software Development Intern"
    }
    template_document = Document(template_file_path)

    for variable_key, variable_value in variables.items():
        for paragraph in template_document.paragraphs:
            replace_text_in_paragraph(paragraph, variable_key, variable_value)

    template_document.save(output_file_path)
    convert(output_file_path, pdf_file_path)


def replace_text_in_paragraph(paragraph, key, value):
    if key in paragraph.text:
        inline = paragraph.runs
        for item in inline:
            if key in item.text:
                item.text = item.text.replace(key, value)


if __name__ == '__main__':
    main()
