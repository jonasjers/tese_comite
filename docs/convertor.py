__author__ = "Edward J. Stembler"
__date__ = "2009-01-09"
__module_name__ = "Converts a batch of Word documents, found in a directory, to text"
__version__ = "1.0"
version_info = (1,0,0)


import sys
import clr
import System
from System.Text import StringBuilder
from System.IO import DirectoryInfo, File, FileInfo, Path, StreamWriter
    
clr.AddReference("Microsoft.Office.Interop.Word")

import Microsoft.Office.Interop.Word as Word


def convert_files(doc_path):

    directory = DirectoryInfo(doc_path)
    files = directory.GetFiles("*.doc")

    for file_info in files:
        text = doc_to_text(Path.Combine(doc_path, file_info.Name))

        stream_writer = File.CreateText(Path.GetFileNameWithoutExtension(file_info.Name) + ".txt")
        stream_writer.Write(text)
        stream_writer.Close()

    return


def doc_to_text(filename):

    word_application = Word.ApplicationClass()
    word_application.visible = False

    document = word_application.Documents.Open(filename)

    result = StringBuilder()

    for p in document.Paragraphs:
        result.Append(clean_text(p.Range.Text))

    document.Close()
    document = None

    word_application.Quit()
    word_application = None

    return result.ToString()


def clean_text(text):

    text = text.replace("\12", "")    # FF
    text = text.replace("\07", "")    # BEL
    text = text.replace("\r", "\r\n") # CR -> CRLF

    return text


test_path = "C:\\test\\"

if __name__ == "__main__":
    if len(sys.argv) == 2:
        convert_files(sys.argv[1])
    else:
        convert_files(test_path)
