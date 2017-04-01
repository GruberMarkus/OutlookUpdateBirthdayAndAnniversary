Option Explicit
forceCScriptExecution

'----------------------------------------------------------
'User adjustable parameters
'----------------------------------------------------------
'Should all visited contacts be logged or only the ones with birthday and/or anniversary dates?
Const OutputOnlyUpdatedContacts = true

'Should all contacts folders be updated or only the default contact folder?
Const DefaultContactFolderOnly = false

'Should hidden folders be updated?
Const UpdateHiddenFolders = false

'When Outlook is not running, the script starts it with the following profile.
'If left empty, you are asked for the profile to use.
Const StartOutlookWithSpecificProfile = ""

'When Outlook is not running, connect to the default profile.
'This overrides option "StartOutlookWithSpecificProfile".
Const StartOutlookWithDefaultProfile = false

'----------------------------------------------------------
'End of user adjustable parameters
'Do not change anything from here on
'----------------------------------------------------------


Dim objOutlook, objNamespace, objParentFolder
Dim oContact, temp, tempBirthday, tempAnniversary
Dim colFolders, objFolder, objSubfolder
Dim str, Folder, WshNetwork, ComputerName, x

Const olFolderContacts = 10

Set WshNetwork = CreateObject("WScript.Network")
ComputerName = WshNetwork.ComputerName

WScript.echo "Connecting to Outlook."
Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
if not StartOutlookWithDefaultProfile = true then objNamespace.Logon StartOutlookWithSpecificProfile

WScript.echo "Searching for contacts with birthday and/or anniversary dates. This may take some time."
WScript.echo
WScript.echo
WScript.echo """Computer"";""Profile"";""Store"";""Folder"";""Last Name"";""First Name"";""File As"";""Birthday ISO"";""Anniversary ISO"""

'For each store (data file) in the profile
For Each str In objNamespace.Stores
	if DefaultContactFolderOnly = false then
		for each objParentFolder in str.getrootfolder.folders
			DoUpdate (objParentFolder)
			GetSubfolders (objParentFolder)
		Next
	end if
	if DefaultContactFolderOnly = true then
		DoUpdate (str.getDefaultFolder(olFolderContacts))
		GetSubfolders (str.getDefaultFolder(olFolderContacts))
	end if
Next

WScript.echo
WScript.echo
WScript.echo "Update complete. You may now close this window."


Sub GetSubfolders(objParentFolder)
	Set colFolders = objParentFolder.Folders
	For Each objFolder In colFolders
		DoUpdate (objNamespace.GetFolderFromID (objFolder.EntryID, objFolder.StoreID))
		GetSubfolders (objNamespace.GetFolderFromID (objFolder.EntryID, objFolder.StoreID))
	Next
End Sub


Sub DoUpdate(Folder)
	'Only allow folders with DefaultItemType of olContactItem
	if folder.DefaultItemType <> 2 then exit sub

	'Do not allow read-only folders
	'Some folder do not have the property set, resulting in an error surviving "on error resume next"
	'therefore this strange construct with in intermediate variable
	x=""
	on error resume next
	x=Folder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10F6000B")
	if x = true then exit sub
	x=""
	err.clear
	on error goto 0

	'Do not allow hidden folders
	'Some folder do not have the property set, resulting in an error surviving "on error resume next"
	'therefore this strange construct with in intermediate variable
	if not UpdateHiddenFolders = true then
		x=""
		on error resume next
		x= Folder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")
		if x = true then exit sub
		x=""
		err.clear
		on error goto 0
	end if

	For Each oContact In Folder.Items
		'Only work with contact class items
		if oContact.class = 40 then
			With oContact
				if OutputOnlyUpdatedContacts = true then wscript.stdout.write chr(13) & string(79," ") & chr(13) & left("Checking " & str.DisplayName & "\" & right(Folder.FolderPath, len(Folder.FolderPath)-len("\\" & str.DisplayName & "\")) & "\" & oContact.FileAs, 79) & chr(13)
				If FormatDateTime(oContact.Birthday, vbShortDate) = FormatDateTime("4501-01-01", vbShortDate) Then tempBirthday = "" Else tempBirthday = Right("0000" & Year(oContact.Birthday), 4) & "-" & Right("00" & Month(oContact.Birthday), 2) & "-" & Right("00" & Day(oContact.Birthday), 2)
				If FormatDateTime(oContact.Anniversary, vbShortDate) = FormatDateTime("4501-01-01", vbShortDate) Then tempAnniversary = "" Else tempAnniversary = Right("0000" & Year(oContact.Anniversary), 4) & "-" & Right("00" & Month(oContact.Anniversary), 2) & "-" & Right("00" & Day(oContact.Anniversary), 2)
				if OutputOnlyUpdatedContacts = false then
					wscript.stdout.write chr(13) & string(79," ") & chr(13)
					WScript.echo """" & ComputerName & """;""" & objNamespace.CurrentProfileName & """;""" & str.displayname & """;""" & right(Folder.FolderPath, len(Folder.FolderPath)-len("\\" & str.displayname & "\")) & """;""" & oContact.LastName & """;""" & oContact.FirstName & """;""" & oContact.FileAs & """;""" & tempBirthday & """;""" & tempAnniversary & """"
				end if
				If Not (tempBirthday = "" And tempAnniversary = "") Then
					if OutputOnlyUpdatedContacts = true then
						wscript.stdout.write chr(13) & string(79," ") & chr(13)
						WScript.echo """" & ComputerName & """;""" & objNamespace.CurrentProfileName & """;""" & str.displayname & """;""" & right(Folder.FolderPath, len(Folder.FolderPath)-len("\\" & str.displayname & "\")) & """;""" & oContact.LastName & """;""" & oContact.FirstName & """;""" & oContact.FileAs & """;""" & tempBirthday & """;""" & tempAnniversary & """"
					end if
					If .Birthday <> FormatDateTime("4501-01-01", vbShortDate) Then
						temp = .Birthday
						.Birthday = FormatDateTime("4501-01-01", vbShortDate)
						.Save
						.Birthday = temp
						.Save
					End If
					If .Anniversary <> FormatDateTime("4501-01-01", vbShortDate) Then
						temp = .Anniversary
						.Anniversary = FormatDateTime("4501-01-01", vbShortDate)
						.Save
						.Anniversary = temp
						.Save
					End If
				End If
			End With
		End If
	Next
	wscript.stdout.write chr(13) & string(79," ") & chr(13)
End Sub


Sub forceCScriptExecution()
	Dim Arg, Str
	If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then
		For Each Arg In WScript.Arguments
			If InStr(Arg, " ") Then Arg = """" & Arg & """"
			Str = Str & " " & Arg
		Next
		CreateObject("WScript.Shell").Run _
			"cmd.exe /k cscript.exe //nologo """ & _
			WScript.ScriptFullName & _
			""" " & Str
		WScript.Quit
	End If
End Sub