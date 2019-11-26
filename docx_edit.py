from docxtpl import DocxTemplate


def create_letter_from_template(context, file_path):

    try:
        doc = DocxTemplate("my_template.docx")
        doc.render(context)
        doc.save(file_path + "\\" + "generated_doc.docx")
        return True
    except Exception as e:
        print("DOC file creation error", e)
        return False
