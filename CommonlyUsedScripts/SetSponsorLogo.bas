'Specify the folder path
dim folderPath as string = "D:\XPression_Projects\_Assets\Logos\"
'Show InputBox and get user input
dim userInput as integer = InputBox("Select a logo type by typing the corresponding number:" & vbCrLf & vbCrLf & "1) Sponsor Logos" & vbCrLf & "2) Tournament Logos" & vbCrLf & "3) Tour Logos" & vbCrLf & "4) Network Logos", "Select a File")
select case userInput
case 1
    folderPath += "Sponsor_Logos\"
case 2
    folderPath += "Tournament_Logos\"
case 3
    folderPath += "Tour_Logos\"
case 4
    folderPath += "Network_Logos\"
end select

' Additional layer for Tournament Logos (userInput=2)
if userInput = 2 then
    if My.Computer.FileSystem.DirectoryExists(folderPath) then
        dim subfolders as System.Collections.ObjectModel.ReadOnlyCollection(Of String)
        subfolders = My.Computer.FileSystem.GetDirectories(folderPath)
        dim promptExtra as string = "Select a subfolder by typing the corresponding number:" & vbCrLf & vbCrLf
        dim j as integer = 1
        for each subPath as string in subfolders
            promptExtra &= j.ToString() & ") " & My.Computer.FileSystem.GetName(subPath) & vbCrLf
            j += 1
        next
        dim userExtra as string = InputBox(promptExtra, "Select a Subfolder")
        dim extraIndex as integer
        if integer.TryParse(userExtra, extraIndex) and extraIndex > 0 and extraIndex <= subfolders.Count then
            folderPath &= My.Computer.FileSystem.GetName(subfolders(extraIndex - 1)) & "\"
            engine.DebugMessage("You selected subfolder: " & folderPath, 0)
        else
            engine.DebugMessage("Invalid subfolder selection. Please enter a valid number.", 0)
            exit sub
        end if
    else
        engine.DebugMessage("Folder does not exist: " & folderPath, 0)
        exit sub
    end if
end if

'Specify the file type of files in the folder
dim fileType as string = ".png"
'Check if the folder exists
if My.Computer.FileSystem.DirectoryExists(folderPath)
    'Get all folders names in the folder
    dim folders as System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    folders = My.Computer.FileSystem.GetDirectories(folderPath)
    'Pagination for Sponsor Logos (userInput=1)
    dim selectedIndex as integer
    if userInput = 1 then
        dim pageSize as integer = 49
        dim currentPage as integer = 0
        dim totalFolders as integer = folders.Count
        dim userInputA as string
				dim endIndex as integer
        do
            dim startIndex as integer = currentPage * pageSize
            endIndex = startIndex + pageSize - 1
            if endIndex >= totalFolders then endIndex = totalFolders - 1
            dim promptA as string = "Select a folder by typing the corresponding number (Page " & (currentPage + 1) & "):" & vbCrLf & vbCrLf
            dim i as integer = 1
            for index as integer = startIndex to endIndex
                promptA &= i.ToString() & ") " & My.Computer.FileSystem.GetName(folders(index)).replace(fileType, "") & vbCrLf
                i += 1
            next
            if endIndex < totalFolders - 1 then
                promptA &= "50) More..." & vbCrLf
            end if
            userInputA = InputBox(promptA, "Select a File")
            if integer.TryParse(userInputA, selectedIndex) and selectedIndex > 0 and selectedIndex <= i then
                if selectedIndex = 50 and endIndex < totalFolders - 1 then
                    currentPage += 1
                else
                    selectedIndex = startIndex + selectedIndex - 1
                    exit do
                end if
            else
                engine.DebugMessage("Invalid selection. Please enter a valid number.", 0)
                exit sub
            end if
        loop while selectedIndex = 50 and endIndex < totalFolders - 1
    else
        'Original folder selection for other cases
        dim promptA as string = "Select a folder by typing the corresponding number:" & vbCrLf & vbCrLf
        dim i as integer = 1
        for each filePath as string in folders
            promptA &= i.ToString() & ") " & My.Computer.FileSystem.GetName(filePath).replace(fileType, "") & vbCrLf
            i += 1
        next
        dim userInputA as string = InputBox(promptA, "Select a File")
        if not integer.TryParse(userInputA, selectedIndex) or selectedIndex <= 0 or selectedIndex > folders.Count then
            engine.DebugMessage("Invalid Selection!", 0)
            exit sub
        end if
        selectedIndex = selectedIndex - 1 ' Adjust for 0-based indexing
    end if

    'Set the file name
    engine.DebugMessage("You selected: " & My.Computer.FileSystem.GetName(folders(selectedIndex)), 0)
    dim sponsorPath as string = folderPath & My.Computer.FileSystem.GetName(folders(selectedIndex)) & "\"
    'Get all file names in the folder
    dim files as System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    files = My.Computer.FileSystem.GetFiles(sponsorPath)
    'Display the file names in InputBox
    dim promptB as string = "Select a logo by typing the corresponding number:" & vbCrLf & vbCrLf
    dim s as integer = 1
    for each filePath as string in files
        promptB &= s.ToString() & ") " & My.Computer.FileSystem.GetName(filePath).replace(fileType, "") & vbCrLf
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
        engine.Sequencer.RegeneratePreview
    else
        engine.DebugMessage("Invalid selection. Please enter a valid number.", 0)
    end if
else
    engine.DebugMessage("Folder does not exist: " & folderPath, 0)
end if