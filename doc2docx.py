from spire.doc import *
from spire.doc.common import *

# Create an object of the Document class\
def parse(sourceFile, destinationFile):
    document = Document()
    document.LoadFromFile(sourceFile)       
    document.SaveToFile(destinationFile, FileFormat.Docx2016)
    document.Close()
    

