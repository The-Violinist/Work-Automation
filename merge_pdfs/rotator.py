import PyPDF2

def rotate():
    print("Rotating the MAV pdf...")
    pdf_in = open('MAV.pdf', 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_in,strict=False)
    pdf_writer = PyPDF2.PdfFileWriter()

    for pagenum in range(pdf_reader.numPages):
        page = pdf_reader.getPage(pagenum)
        page.rotateClockwise(90)
        pdf_writer.addPage(page)

    pdf_out = open('mav_rotated.pdf', 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()
