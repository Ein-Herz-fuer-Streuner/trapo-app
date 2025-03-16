import os

folder = "./data/traces"

pdfs = []
for filename in os.listdir(folder):
    if filename.endswith(".pdf"):
        full_path = os.path.join(folder, filename)
        pdfs.append(full_path)
    else:
        continue

for old in pdfs:
    dir, file = os.path.split(old)
    new = file.split("_",1)[0]+".pdf"
    new = new.replace(".pdf.pdf", ".pdf")
    new_path = os.path.join(dir, new)
    os.rename(old, new_path)