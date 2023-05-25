'-renames-shortcuts------------------------------------------------------------


Dim maxEp
Dim remainingEpisodes
Dim percentage
Dim finalNameExactlyFolderName
Dim txtUnderscore
Dim exactFolderName
Dim aa
Dim extension
Dim seasonNum
Dim nowant

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


	'goes in ech subfolder and calculates percentage
	For Each oneSeriesFolder in seriesFolders
		If (oneSeriesFolder.Name <> "_mkvs") Then
			
			'get files inside each sub folder in the target folder
			Set oneSeriesFolderPath = objFSO.GetFolder(objStartFolder & "\" & oneSeriesFolder.Name)
			
			Set episodeFiles = oneSeriesFolderPath.Files

			'Wscript.Echo "Name of the Folder :" & oneSeriesFolder.Name

			' -------------------------------------------------
				'rename shortcuts

				'get each file[shortcut] in target folder
				For Each oneShortcut in shortcutFiles
					
					'splits the extension and percentage bracket if any then get the "name of foler 03"
					a=Split(oneShortcut.Name, ".")
					extension = a(1)
					'check if file is a shortcut and splits it at ' [48%] (5)'
					If extension = "lnk" Then
						aa = a(0)
						If Right(aa, 2) = "%]" OR Right(aa, 4) = "TED]" Then
							a=Split(aa, " [")
						End If
						scname = a(0)
						'WScript.Echo scname
						
						If Left (scname, 1) = "(" Then

						b = Split(scname, ") ")
						
						scname = b(1)

						'WScript.Echo scname						
						End If
						'Wscript.Echo "shortcut name with number :" & scname
						'teriffier 02
					
						'removes season number from string
						nowant = Right(scname, 3)
						isItNumber = Left(nowant, 2)
						If isItNumber = " 0" OR isItNumber = " 1" OR isItNumber = " 2" OR isItNumber = " 3" OR isItNumber = " 4" Then
							finalNameExactlyFolderName = Split(scname, nowant, -1, 1)
							exactFolderName = finalNameExactlyFolderName(0)
						Else 
							exactFolderName = scname
						End If
						'Wscript.Echo "only shortcut name like folder:" & exactFolderName
				
				
						'matches the folder name with shortcut
						If LCase(exactFolderName) = LCase(oneSeriesFolder.Name) Then
							'Wscript.Echo a(0) & " Matched"
							'renames the shortcut and puts the percentage of episodes watched in the name



							' -------------------------------------------------------------
				
							seasonNum = Right(nowant, 2)
							'to get last episode number
							remainingEpisodes = 0
							'get each file episode number if "03. Ep Name.mkv" then returns 03
							For Each oneEpisode in episodeFiles
								if InStr(1, oneEpisode, " - ", 0) > 0 AND InStr(1, oneEpisode, "x", 0) > 0 Then
									a = Split(oneEpisode.Name, " - ")
									b = Split(a(0), "x")
									' WScript.Echo aa
									currentEpSeason = b(0)
									currentEp = b(1)

									'Wscript.Echo "season " & currentEpSeason & " episode " & currentEp & " folder season " & seasonNum

									
									if currentEpSeason = seasonNum Then
										maxEp= currentEp

										'Wscript.Echo maxEp

										remainingEpisodes = remainingEpisodes + 1
										'Wscript.Echo remainingEpisodes
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
								'Wscript.Echo "percentage of completed eps:" & percent
							Else 
								percent = "[SEASON COMPLETED]"
							End If

							' -------------------------------------------------------------




							Dim oldSCName
							oldSCName = oneShortcut.Name
							Dim newSCName
							newSCName =  "(" & remainingEpisodes & ") " & scname & " " & percent & ".lnk"
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