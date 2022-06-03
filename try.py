from fpdf import FPDF
import excelrd as ex


class PDF(FPDF):
    def header(self):
        # Logo
        self.image('logo.jpg', 10, 8, 33)
        # Arial bold 15
        self.set_font('Arial', 'B', 15)
        # Move to the right
        self.cell(50)
        # Title
        self.cell(120, 20, 'Swarrnim Startup And Innovation University', 1, 0, 'C')
        # Line break
        self.ln(20)

    # Page footer
    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-15)
        # Arial italic 8
        self.set_font('Arial', 'I', 8)
        # Page number
        self.cell(0, 10, 'Page ' + str(self.page_no()) + '/{nb}', 0, 0, 'C')


loc = "test.xlsx"

wb = ex.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
n = sheet.nrows
c = sheet.ncols
print("Number of rows are ", n)
for i in range(1, n):
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_font('Times', '', 20)
    pdf.ln(20)
    title = sheet.row_values(0)
    line = sheet.row_values(i)
    print(line)
    for j in range(sheet.ncols):
        pdf.cell(0, 0, title[j], 0, 1)
        pdf.cell(110)
        pdf.cell(0, 0, "Blank ", 0, 0)
        pdf.ln(20)
    x = "try.pdf"
    pdf.output(x, 'F')
