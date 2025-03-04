import camelot

file = "data/petoffice.pdf"
#print(f"Using camelot v{camelot.__version__}.")
tables = camelot.read_pdf(file, pages="all", flavor="lattice", backend="pdfium", line_scale=20)
for tab in tables:
    print(tab.df)