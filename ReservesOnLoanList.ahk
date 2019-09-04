;Start;
final_list = ReservesOnLoan.docx
FileDelete %final_list%

IfNotExist, Template.docx
{
	msgbox Cannot find Template.docx
	exit
}

;Configurations
IniRead, path, config.ini, xls_path, path

;Get input file
FileSelectFile, xlsFile,,%path%, Select File, *.xls*

;Check for input file or cancel to exit
If xlsFile =
{
	exit
}

;Status
Progress, zh0 fs12, Generating List...One Moment...,,Status

;Open DOC file
template = %A_ScriptDir%\Template.docx
saveFile = %A_ScriptDir%\%final_list%
wrd := ComObjCreate("Word.Application")
wrd.Visible := False

;Perform Mail Merge
doc := wrd.Documents.Open(template)
doc.MailMerge.MainDocumentType := 3 ;Mail merge type "directory"
doc.MailMerge.OpenDataSource(xlsFile,,,,,,,,,,,,,"SELECT * FROM [results$]")
doc.MailMerge.Execute

;Add header row
wrd.Selection.InsertRowsAbove(1)
wrd.Selection.Tables(1).Rows(1).Height := 30
wrd.Selection.Cells.VerticalAlignment := 1
wrd.Selection.ParagraphFormat.Alignment := 1
wrd.Selection.Shading.BackgroundPatternColor := -587137025
wrd.Selection.Font.Italic := False
wrd.Selection.Font.Bold := True
wrd.Selection.TypeText("Title")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Barcode")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Due Date")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("Course Code")
wrd.Selection.MoveRight(1)
wrd.Selection.TypeText("On Shelf?")
wrd.Selection.Rows.HeadingFormat := 9999998 ;Set header for each page
    
;Save and quit DOC file
wrd.ActiveDocument.SaveAs(saveFile)
wrd.DisplayAlerts := False
doc.Close
wrd.Quit

;Finish
IfNotExist, %final_list%
{
	msgbox Cannot find %final_list%
	exit
}
FileDelete %xlsFile%
run winword.exe %final_list%
