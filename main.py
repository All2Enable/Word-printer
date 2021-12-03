from win32com import client
import time

word = client.Dispatch("Word.Application")


def printWordDocument(Lab1):
    word.Documents.Open(Lab1)
    time.sleep(2)
    word.ActiveDocument.Close()


word.Quit()
