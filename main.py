from win32com import client
import time

word = client.Dispatch("Word.Application")


def printWordDocument(filename):
    word.Documents.Open(filename)
    word.ActiveDocument.PrintOut()
    time.sleep(2)
    word.ActiveDocument.Close()

printWordDocument()
word.Quit()
