isTest = True
args = WScript.Arguments.Count

If args <> 4 Then
	WScript.Echo "Incorrect number of arguments specified."
	WScript.Echo "Usage: cscript mailer.vbs <path to XLS> <worksheet> <office location> <PDF archive location>"
	WScript.Echo "    office location: BLR | MAA | PNQ | DEL"
	WScript.Quit
End If

'Constants
senderName = "GodFather"
senderEmail = "khushroc@thoughtworks.com"
testEmail = "khushroo.cooper@gmail.com"

' Variables
xlsPath = WScript.Arguments.Item(0)
worksheetName = WScript.Arguments.Item(1)
officeLocation = WScript.Arguments.Item(2)
pdfLocation = WScript.Arguments.Item(3)

Select Case officeLocation
	Case "PNQ"
		cCellNumber = 2
		cName = 3
		cSubTotal = 4
		cTax = 5
		cTotal = 6
		cEligibility = 7
		cExcess = 8
		cEmpNumber = 9
		cEmail = 10
		cFileName = 1
	Case "DEL"
		cEmpNumber = 1
		cName = 2
		cCellNumber = 3
		cSubTotal = 6
		cTax = 7
		cTotal = 8
		cEligibility = 9
		cExcess = 10
		cEmail = 11
		cFileName = 3
	Case Else
		WScript.Echo "Incorrect office location specified."
		WScript.Quit
End Select

WScript.Echo "> File: " & xlsPath
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = 0
objExcel.Workbooks.open xlsPath, false, true

WScript.Echo "> Worksheet: " & worksheetName
Set myWorksheet = objExcel.Sheets(worksheetName)
usedColumnsCount = myWorksheet.UsedRange.Columns.Count
usedRowsCount = myWorksheet.UsedRange.Rows.Count

Set Cells = myWorksheet.Cells

' Get CWD
Set objFSO = CreateObject("Scripting.FileSystemObject")
cwd = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Templates"
Set objFSO = Nothing

textTemplate = GetTemplate(cwd & "Template.txt")
htmlTemplate = GetTemplate(cwd & "Template.html")

myPassword = InputBox("Security Check", "Enter your GMail password")

' Loop through each row in the worksheet. Top row assumed Header
For row = 2 to (usedRowsCount-1)
	myCellNumber = Cells(row, cCellNumber)
	myName = Cells(row, cName)
	myEmpNumber = Cells(row, cEmpNumber)
	mySubTotal = FormatNumber(Cells(row, cSubTotal),2)
	myTax = FormatNumber(Cells(row, cTax),2)
	myTotal = FormatNumber(Cells(row, cTotal),2)
	myEligibility = Cells(row, cEligibility)
	myExcess = FormatNumber(Cells(row, cExcess), 2)
	myEmail = Cells(row, cEmail)

	myPDF = Cells(row, cFileName) & ".pdf"

	If isTest = True Then 
		myEmail = testEmail
		myPDF = "test.pdf"
	End If
	
	myText = GenerateEmailText(False)	
	myHTML = GenerateEmailText(True)
	
	WScript.Echo myEmpNumber & ":" & myName & ":" & myEmail
	
	SendEmail
	
	WScript.Echo "------------------------------------------"
	
	If isTest = True Then Exit For
Next

' We are done with the current worksheet, release the memory
Set myWorksheet = Nothing

objExcel.Workbooks(1).Close
objExcel.Quit

Set myWorksheet = Nothing
Set objExcel = Nothing

Function GetTemplate(templateFileName)
	WScript.Echo "> Reading template: " & templateFileName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile (templateFileName, 1)
	strText = objTextFile.ReadAll
	objTextFile.Close
	
	Set ojbTextFile = Nothing
	Set objFSO = Nothing

	GetTemplate = strText
End Function

Function GenerateEmailText(isHTML)
	If isHTML = True Then
		verbiage = htmlTemplate
	Else
		verbiage = textTemplate
	End If
	
	verbiage = Replace(verbiage, "$NAME$", myName)
	verbiage = Replace(verbiage, "$MONTH$", worksheetName)
	
	verbiage = Replace(verbiage, "$EMPNUMBER$", myEmpNumber)
	verbiage = Replace(verbiage, "$NAME$", myName)
	verbiage = Replace(verbiage, "$NUMBER$", myCellNumber )
	verbiage = Replace(verbiage, "$SUBTOTAL$", mySubTotal)
	verbiage = Replace(verbiage, "$TAX$", myTax)
	verbiage = Replace(verbiage, "$TOTAL$", myTotal)
	verbiage = Replace(verbiage, "$ELIGIBILITY$", myEligibility)
	verbiage = Replace(verbiage, "$EXCESS$", myExcess)
	verbiage = Replace(verbiage, "$PDF$", myPDF)

	GenerateEmailText = verbiage
End Function

Function SendEmail
	Set objEmail      = CreateObject("CDO.Message")
	objEmail.Subject  = "Your mobile telephone bill for " & worksheetName
	objEmail.Sender   = senderEmail
	objEmail.From     = """" & senderName & """ <" & senderEmail & ">"
	objEmail.To       = myEmail
	objEmail.TextBody = myText
	objEmail.HTMLBody = myHTML
	objEmail.AddAttachment pdfLocation & myPDF
	
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = senderEmail
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = myPassword
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.googlemail.com"
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
 
	objEmail.Configuration.Fields.Update	
	objEmail.Send

	WScript.Echo "Mail Sent."
End Function