import tabula
import pandas as pd

#declare the path of your file
file_path = "C:\\Users\\Karim\\Downloads\\test.pdf"
#Convert your file
df = tabula.read_pdf(file_path)