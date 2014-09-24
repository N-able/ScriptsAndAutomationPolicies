' VB Script Document
option explicit
dim strLogFile, FSO, TextStream
Const FOR_APPENDING = 8 'For more information on this variable, check out http://msdn.microsoft.com/en-us/library/314cz14s%28VS.85%29.aspx


'Configure what file you want to create and write to
strLogFile = "C:\MyTextFile.txt"



Set FSO = CreateObject("Scripting.FileSystemObject")
Set TextStream = FSO.OpenTextFile(strLogFile, FOR_APPENDING, True)
TextStream.WriteLine "Look at me - I wrote text into a file! Huzzah!"
TextStream.Close
Set FSO = Nothing