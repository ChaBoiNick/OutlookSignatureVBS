On Error Resume Next

'References
'All objuser.XXXX and there counterparts in AD 
'https://ss64.com/vb/syntax-userinfo.html

'Changelog 
'v1.00 2/5/18
'Script will generate htm txt and rich docs for a ms signature using the data in active directory
'Generated in local pc and sets as send and recieve default signature
'Adds any banner image named new.jpg in signature with set dimensions

'v1.01 3/5/18
'Created if statements for mobile info so skips if no data present
'Added comments

'v1.02 7/5/18
'Expanded if statement for mobile to all contact information
'Some slight formatting changes

'v1.03 


'Current projects
'If statement for second picture (badge only-badge.jpg) to add inline next to new.jpg dimensions assumed square
'Create external script to update image
'Select image from browse menu and input desired dimensions to be updated within this script

Set objSysInfo = CreateObject("ADSystemInfo")
Set WshShell = CreateObject("WScript.Shell")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Title
strCred = objUser.info
strStreet = objUser.StreetAddress
strState = objUser.st
strLocation = objUser.l
strPostCode = objUser.PostalCode
strPhone = objUser.TelephoneNumber
strDirect = objUser.ipPhone
strMobile = objUser.Mobile
strEmail = objUser.mail
strWebsite = objUser.wWWHomePage
strOffice = objUser.physicalDeliveryOfficeName

'Creates word application for formatting
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

'Signature Font 
objSelection.Font.Name = "Verdana"
objSelection.Font.Size = 10 'Carries over unless specified again elsewhere

'Salutation
objSelection.font.color = rgb(0,0,0)
objSelection.TypeText "Regards,"

'Line break
'objSelection.TypeText Chr(11)
objSelection.TypeParagraph()

'Username line
objSelection.Font.Size = 12
objSelection.Font.Bold = true
if (strCred) Then objSelection.TypeText strName & ", " & strCred Else objSelection.TypeText strName
objSelection.Font.Bold = false

'Job title line
objSelection.Font.Size = 10
objSelection.TypeParagraph()
objSelection.ParagraphFormat.LineSpacing = 16
objSelection.TypeText strTitle
objSelection.TypeText Chr(11)

'Location line
objSelection.Font.Bold = true
objSelection.font.color = rgb(210,73,42)
objSelection.TypeText strOffice & " Office " & "| FLOTH Sustainable Building Consultants"
objSelection.Font.Bold = False
objSelection.TypeText Chr(11)

'Address line
objSelection.Font.Size = 9
objSelection.font.color = rgb(0,0,0)
objSelection.TypeText strStreet & ", " & strLocation & ", " & strState & ", " & strPostCode
objSelection.TypeText Chr(11)

'Contact line
'Formatted to print results horizontally - to print vertically add objSelection.TypeText Chr(11) in between each object
objSelection.Font.Size = 8
objSelection.font.color = rgb(0,0,0)

'If the data is not present in the AD it will not print anything and move on to the next field.
If Not IsEmpty(strPhone) Then
    objselection.typetext "P: " & strPhone
End If

If Not IsEmpty(strDirect) Then
    objselection.typetext " | D: " & strDirect
End If

If Not IsEmpty(strmobile) Then
    objselection.typetext " | M: " & strMobile
End If

If Not IsEmpty(strEmail) Then
    objselection.typetext " | E: " & strEmail
End If

If Not IsEmpty(strWebsite) Then
    objselection.typetext " | W: " & strWebsite
End If

objSelection.TypeText Chr(11)

' If statement to hyperlink website 
' Don't really need this as most email clients auto format the email and website to hyperlinks
' if strWebsite then
' Set objLink = objSelection.Hyperlinks.Add(objselection.Range,strWebsite)
	' objLink.Range.Font.Name = "Verdana"
	' objLink.Range.Font.Size = 8
	' objLink.Range.Font.Bold = false
' end if
' objSelection.TypeText Chr(11)

'Image description or disclaimer
objSelection.Font.Size = 9
objSelection.Font.Bold = true
objSelection.font.color = rgb(0,187,0)
objSelection.TypeText "Winner of the 2017 Brisbane Lord Mayors Business Awards for Sustainability in Business, awarded to Floth for 69 Robertson Street, Fortitude Valley."
objSelection.Font.Bold = false
objSelection.TypeText Chr(11)

'New signature image adding - Place script and file in NETLOGON and adjust image file path
Set shp = objSelection.InlineShapes.AddPicture("\\flsvr03\Software\_New Machine Install\SIGNATURES New\Dev\test.jpg")
shp.LockAspectRatio = msoFalse
shp.Width = 456
shp.Height = 86

'Can make an if statement for if there is a badge signature instead of a banner.

'Code for multuple departments with different signature images
' If (objUser.Department = "Department NAME") Then 
             ' objSelection.InlineShapes.AddPicture("\LMBA_Landscape_DarkGreen_668x126.jpg") 
 
 
' ElseIf (objUser.Department = "Department NAME") Then 
        ' objSelection.InlineShapes.AddPicture("\LMBA_Landscape_DarkGreen_668x126.jpg") 
 
' Else 
        ' objSelection.InlineShapes.AddPicture("\LMBA_Landscape_DarkGreen_668x126.jpg") 
 
' End If 

Set objSelection = objDoc.Range()

objSignatureEntries.Add "EmailSignature", objSelection 
objSignatureObject.NewMessageSignature = "EmailSignature" 
objSignatureObject.ReplyMessageSignature = "EmailSignature" 

objDoc.Saved = True
objWord.Quit