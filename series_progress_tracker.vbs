Dim maxEp
Dim remainingEpisodes
Dim percentage
Dim finalNameExactlyFolderName
Dim exactFolderName
Dim extension
Dim seasonNum

'target path
objStartFolder = "path\to\your\seriesFolder" 'example: 'G:\The  W. A. F. O. L. T. S.  Folder\W. A. F. O. S'

'interact with files
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(objStartFolder)

'get files[shortcuts] and folders in the target folder
Set seriesFolders = objFolder.Subfolders
Set shortcutFiles = objFolder.Files

'runs every 5 seconds
Do
	'goes in each subfolder and calculates percentage
	For Each oneSeriesFolder in seriesFolders
		If (oneSeriesFolder.Name <> "_mkvs") Then
			'get files inside each subfolder in the target folder
			Set oneSeriesFolderPath = objFSO.GetFolder(objStartFolder & "\" & oneSeriesFolder.Name)
			Set episodeFiles = oneSeriesFolderPath.Files
			
			'rename shortcuts
			For Each oneShortcut in shortcutFiles
				'splits the extension and percentage bracket if any, then gets the "name of folder 03"
				a = Split(oneShortcut.Name, ".")
				extension = a(1)
				
				'check if file is a shortcut and splits it at ' [48%] (5)'
				If extension = "lnk" Then
					aa = a(0)
					If Right(aa, 2) = "%]" Or Right(aa, 4) = "TED]" Then
						a = Split(aa, " [")
					End If
					
					scname = a(0)
					
					If Left(scname, 1) = "(" Then
						b = Split(scname, ") ")
						scname = b(1)
					End If
					
					'removes season number from string
					nowant = Right(scname, 3)
					isItNumber = Left(nowant, 2)
					
					If isItNumber = " 0" Or isItNumber = " 1" Or isItNumber = " 2" Or isItNumber = " 3" Or isItNumber = " 4" Then
						finalNameExactlyFolderName = Split(scname, nowant, -1, 1)
						exactFolderName = finalNameExactlyFolderName(0)
					Else 
						exactFolderName = scname
					End If
					
					'matches the folder name with shortcut
					If LCase(exactFolderName) = LCase(oneSeriesFolder.Name) Then
						'renames the shortcut and puts the percentage of episodes watched in the name
						seasonNum = Right(nowant, 2)
						remainingEpisodes = 0
						
						'get each file episode number if "03. Ep Name.mkv" then returns 03
						For Each oneEpisode in episodeFiles
							If InStr(1, oneEpisode, " - ", 0) > 0 And InStr(1, oneEpisode, "x", 0) > 0 Then
								a = Split(oneEpisode.Name, " - ")
								b = Split(a(0), "x")
								
								currentEpSeason = b(0)
								currentEp = b(1)
								
								If currentEpSeason = seasonNum Then
									maxEp = currentEp
									 
									remainingEpisodes = remainingEpisodes + 1
								End If
							End If
						Next
						
						'handles divide by zero error
						If (maxEp > 0) Then
							percentage = (remainingEpisodes / maxEp * (-100)) + 100
							If percentage <> 100 Then
								percent = "[" & Round(percentage) & "%]"
							Else
								percent = "[SEASON COMPLETED]"
							End If
						Else 
							percent = "[SEASON COMPLETED]"
						End If
						
						Dim oldSCName
						oldSCName = oneShortcut.Name
						Dim newSCName
						newSCName = "(" & remainingEpisodes & ") " & scname & " " & percent & ".lnk"
						If oldSCName <> newSCName Then
							objFso.MoveFile objStartFolder & "\" & oldSCName , objStartFolder & "\" & newSCName
						End If
					End If
				End If
			Next
		End If
	Next
	
	'runs every 10 seconds
	WScript.Sleep(10000)
Loop
