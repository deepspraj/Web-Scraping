import vocabulary
import pandas as pd

print("Make sure that the path(location) must have two backslashes(\\\\) after every folder (e.g: D:\\\\vocabulary.docx). Else file will be saved at default path i.e C:\\vocabulary.docx")

locationRetrieve = input("Location of data file :\n")
locationStore = input("Location to store the docx file (editable) :\n")

df = pd.read_csv(locationRetrieve)

print("Your desired file will be saved at :" + locationStore)


for i in range (len(df)):
    instant = vocabulary.meaningGet(df.at[i,'Words'], locationStore)
    instant.meaning()
    