'Specify the folder path
dim folderPath as string = "D:\XPression_Projects\_Assets\Logos\"

'Show InputBox and get user input
dim userInput as integer = InputBox("Select a logo type by typing the corresponding number:" & vbCrLf & vbCrLf & "1) Sponsors" & vbCrLf & "2) Tournaments", "Select a File")

select case userInput
	case 1
		folderPath += "Sponsor_Logos\"
	case 2
		folderPath += "Tournament_Logos\"
end select

'Specify the file type of files in the folder
dim fileType as string = ".png"

'Check if the folder exists
if My.Computer.FileSystem.DirectoryExists(folderPath)
	'Get all folders names in the folder
	dim folders as System.Collections.ObjectModel.ReadOnlyCollection(Of String)

	folders = My.Computer.FileSystem.GetDirectories(folderPath)

	'Display the file names in InputBox
	dim promptA as string = "Select a sponsor by typing the corresponding number:" & vbCrLf & vbCrLf

	dim i as integer = 1

	for each filePath as string in folders
		promptA &= i.ToString() & ") " & My.Computer.FileSystem.GetName(filePath).replace(fileType,"") & vbCrLf
		i += 1
	next

	'Show InputBox and get user input
	dim userInputA as string = InputBox(promptA, "Select a File")

	'Convert the user's input to a valid selection
	dim selectedIndex as integer

	if integer.TryParse(userInputA, selectedIndex) and selectedIndex > 0 and selectedIndex <= folders.Count
		'Set the file name
		engine.DebugMessage("You selected: " & My.Computer.FileSystem.GetName(folders(selectedIndex - 1)), 0)

		dim sponsorPath as string = folderPath & My.Computer.FileSystem.GetName(folders(selectedIndex - 1)) & "\"

		'Get all file names in the folder
		dim files as System.Collections.ObjectModel.ReadOnlyCollection(Of String)

		files = My.Computer.FileSystem.GetFiles(sponsorPath)

		'Display the file names in InputBox
		dim promptB as string = "Select a logo by typing the corresponding number:" & vbCrLf & vbCrLf

		dim s as integer = 1

		for each filePath as string in files
			promptB &= s.ToString() & ") " & My.Computer.FileSystem.GetName(filePath).replace(fileType,"") & vbCrLf
			s += 1
		next

		'Show InputBox and get user input
		dim userInputB as string = InputBox(promptB, "Select a File")

		'Convert the user's input to a valid selection
		dim logoIndex as integer

		if integer.TryParse(userInputB, logoIndex) and logoIndex > 0 and logoIndex <= files.Count
			'Set the file name
			dim logoFile as string = sponsorPath & My.Computer.FileSystem.GetName(files(logoIndex - 1))
			dim blurFile as string = sponsorPath & "Blurred\" & My.Computer.FileSystem.GetName(files(logoIndex - 1))
			engine.DebugMessage("You selected: " & logoFile, 0)

			dim scene as xpScene
			dim takeItem as xpTakeItem
			dim published as xpPublishedObject
			dim mat as xpMaterial
			dim sha as xpBaseShader

			if engine.Sequencer.GetFocusedTakeItem(takeItem)
				if takeItem.Name.Contains("Leaderboard")
					engine.GetMaterialByName("FS_Sponsor_Logo", mat)
					mat.GetShader(0, sha)
					sha.SetFileName(logoFile)
					mat.UpdateThumbnail
				else
					if takeItem.GetPublishedObjectByName("<<SPONSOR SELECT>>", published)
						published.SetPropertyString(0, logoFile)
					else if takeItem.GetPublishedObjectByName("<< SPONSOR SELECT >>", published)
						published.SetPropertyString(0, logoFile)
					end if
					
					if takeItem.GetPublishedObjectByName("<<SPONSOR BLUR>>", published)
						published.SetPropertyString(0, blurFile)
					else if takeItem.GetPublishedObjectByName("<< SPONSOR BLUR >>", published)
						published.SetPropertyString(0, blurFile)
					end if
				end if
			end if
		end if

		engine.Sequencer.RegeneratePreview
	else
		engine.DebugMessage("Invalid selection. Please enter a valid number.", 0)
	end if
else
	engine.DebugMessage("Invalid Selection!", 0)
end if