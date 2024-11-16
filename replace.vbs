Set objWord = CreateObject("Word.Application")
objWord.Visible = False
Set objDoc = objWord.Documents.Add()
objWord.VBE.ActiveVBProject.VBComponents.Import "C:\Users\Meiram\Documents\orders\mymodule.bas"
objWord.Run "ReplaceVariableInMultiFiles"
objDoc.Close False
objWord.Quit