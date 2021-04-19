from PyPDF2 import PdfFileWriter, PdfFileReader
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject

def set_need_appearances_writer(writer):

    try:
        catalog = writer._root_object
        if "/AcroForm" not in catalog:
            writer._root_object.update({
                NameObject("/AcroForm"): IndirectObject(len(writer._objects), 0, writer)})

        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        return writer

    except Exception as e:
        print('set_need_appearances_writer() catch : ', repr(e))
        return writer

def create_pdf(infile, outfile, field_dictionary):

    inputStream = open(infile, "rb")
    pdf_reader = PdfFileReader(inputStream, strict=False)
    if "/AcroForm" in pdf_reader.trailer["/Root"]:
        pdf_reader.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})

    pdf_writer = PdfFileWriter()
    set_need_appearances_writer(pdf_writer)
    if "/AcroForm" in pdf_writer._root_object:
        pdf_writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)})

    pdf_writer.addPage(pdf_reader.getPage(0))
    pdf_writer.updatePageFormFieldValues(pdf_writer.getPage(0), field_dictionary)

    outputStream = open(outfile, "wb")
    pdf_writer.write(outputStream)
    outputStream.close()
    inputStream.close()

def get_headers(infile):
    inputStream = open(infile, "rb")
    pdf_reader = PdfFileReader(inputStream, strict=False)
    fields = pdf_reader.getFields().keys()
    print('Headers needed for data file: {}'.format(list(fields)))

