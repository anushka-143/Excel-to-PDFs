from fpdf import FPDF
import excelrd as ex
import pandas as pd


class PDF(FPDF):
    def header(self):
        # Logo
        self.image('logo.jpg', 10, 8, 33)
        # Arial bold 15
        self.set_font('Arial', 'B', 15)
        # Move to the right
        self.cell(50)
        # title
        head = 'Swarrnim Startup And Innovation University'
        self.cell(120, 20, head, 1, 0, 'C')
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


# location(name) of the file
a = pd.read_excel('test.xlsx')
loc = "test.xlsx"
# accessing workbook
wb = ex.open_workbook(loc)
# accessing worksheet
sheet = wb.sheet_by_index(0)
# accessing particular cell value
sheet.cell_value(0, 0)
# number of rows and cols in a sheet
n = sheet.nrows
c = sheet.ncols
print("Number of rows are ", n)

# Loop for creating multiple Pdfs
for i in range(1, n):
    # Making an object of PDF class
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.add_page()
    # setting font and size
    pdf.set_font('Times', '', 20)
    # line break
    pdf.ln(20)
    # calling first row
    title = sheet.row_values(0)
    # calling specific rows by 'i'
    line = sheet.row_values(i)
    # print(line)
    # loop for collecting data from each cell in specified row
    for j in range(sheet.ncols):
        # writing key into the pdf
        pdf.cell(0, 0, title[j], 0, 1)
        # giving space in same line
        pdf.cell(110)
        # writing value into the pdf
        pdf.cell(0, 0, " " + str(line[j]), 0, 0)
        # Next line
        pdf.ln(20)
    # Storing Directory + Filename + extension
    # taking unique values for filename!
    # you can replace,
    # str(sheet.cell_value(i, 5)) + str(sheet.cell_value(i, 8))
    # with 'your own text' for file name!
    x = "pdfs\\" + str(sheet.cell_value(i, 5)) + str(sheet.cell_value(i, 8)) + ".pdf"
    # saving pdf
    pdf.output(x, 'F')
