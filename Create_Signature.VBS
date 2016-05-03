On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName

Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strFirstName = objuser.givenName
StrLastName = objuser.sn
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strAddress = objuser.streetAddress
strCity = objuser.l
strState = objuser.st
strZip = objuser.postalCode
strPhone = objUser.extensionAttribute10
strMobilePhone = objUser.mobile
strMail = objuser.mail
strLocation = objuser.physicalDeliveryOfficeName
strPager = objuser.Pager
strWeb = objuser.wWWHomePage
strLogo = "\\domain\netlogon\logo_n1.jpg" 'Location of Logo file
strDisc = ""


Set objWshNet = CreateObject("Wscript.Network")
strCN = lcase(objWshNet.ComputerName)
strCU = lcase(objWshNet.UserName)


Set objWord = CreateObject("Word.Application")
objWord.Visible = False

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objRange = objDoc.Range()

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
Set colShapes = objDoc.Shapes


objSelection.Font.Name = "Arial"
objSelection.Font.Color = RGB(0,0,0)

objDoc.Tables.Add objRange, 1, 2
Set objTable = objDoc.Tables(1)
	objTable.Cell(1, 1).select
	objTable.Columns(1).Width = objWord.InchesToPoints(2)
		colShapes.AddPicture(strLogo)
	objTable.Cell(1, 2).select
		objSelection.Font.Bold = True
		objSelection.Font.Size = "10"
		objSelection.Font.Name = "Arial"
		objSelection.TypeText strName
		objSelection.TypeText(Chr(11))
		objSelection.Font.Size = "8"
		objSelection.TypeText strDepartment
		objSelection.TypeText(Chr(11))
		objSelection.TypeText(Chr(11))
		objSelection.Font.Size = "10"
		objSelection.TypeText "COMPANY SA" 'Type the name of your company
		objSelection.Font.Bold = false
		objSelection.Font.Size = "8"
		objSelection.Font.Name = "Calibri"
		objSelection.TypeText(Chr(11))
		objSelection.TypeText(Chr(11))
		objSelection.TypeText strPager & strPhone
		objSelection.TypeText(Chr(11))
		'---------------------------------------------------------------------------------
		' Define fonte e-mail
		'---------------------------------------------------------------------------------
		objSelection.TypeText " "
		With objSelection.Hyperlinks.Add (objSelection.range, "mailto:" & strMail, , , strMail)
			.Range.Font.Size = "8"
			.Range.Font.Name = "Calibri"
			.Range.Bold = False
			.Range.SubscribeTo = False
			.Range.Font.Color = RGB(0, 0, 0)
		End With
		objSelection.TypeText(Chr(11))
		
		'---------------------------------------------------------------------------------
		' Define SITE
		'---------------------------------------------------------------------------------
		objSelection.TypeText " "
		With objSelection.Hyperlinks.Add (objSelection.range, "http://siteofyourcompany.com", , , "http://siteofyourcompany.com")
			.Range.Font.Size = "8"
			.Range.Font.Name = "Calibri"
			.Range.Bold = False
			.Range.SubscribeTo = False
			.Range.Font.Color = RGB(0, 0, 0)
		End With
		objSelection.TypeText(Chr(11))
	
		'----------------------------------------------------------------------------------

objSelection.EndKey END_OF_STORY

objSignatureEntries.Add "Signature", objRange
objSignatureObject.NewMessageSignature = "Signature"
objSignatureObject.ReplyMessageSignature = "Signature"

objDoc.Saved = True
objWord.Quit
'Wscript.echo "OK"
