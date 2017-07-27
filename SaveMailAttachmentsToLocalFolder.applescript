using terms from application "Mail"
	on perform mail action with messages theMessages for rule theRule
		
		-- set up the attachment folder path
		tell application "Finder"
			--set attachmentsFolder to (path to documents folder as text) as text			
			set folderName to "Box"
			--set homePath to (path to home folder as text) as text
			set documentsPath to (path to documents folder as text) as text
			set attachmentsFolder to (documentsPath & folderName) as text
		end tell
		
		tell application "Mail"
			repeat with eachMessage in theMessages
				--set the sub folder for the attachments to the mail's subject.
				set subFolder to (subject of eachMessage)
				
				set {year:y, month:m, day:d, hours:h, minutes:min, seconds:sec} to eachMessage's date received
				
				set timeStamp to (y & "-" & my pad(m as integer) & "-" & my pad(d) & "_" & my pad(h) & "-" & my pad(min) & "-" & my pad(sec) & "_") as string
				
				set attachCount to count of (mail attachments of eachMessage)
				
				if attachCount is not equal to 0 then
					-- use the unix /bin/test command to test if the timeStamp folder  exists. if not then create it and any intermediate directories as required
					if (do shell script "/bin/test -e " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder) & " ; echo $?") is "1" then
						-- 1 is false
						do shell script "/bin/mkdir -p " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder)
					end if
					try
						-- Save the attachment
						repeat with theAttachment in eachMessage's mail attachments
							set originalName to name of theAttachment
							set newFileName to (timeStamp & originalName) as string
							set savePath to attachmentsFolder & ":" & subFolder & ":" & newFileName
							try
								save theAttachment in file (savePath)
							end try
						end repeat
						--on error msg
						--display dialog msg
					end try
				end if
			end repeat
		end tell
	end perform mail action with messages
end using terms from
on pad(n)
	return text -2 thru -1 of ("0" & n)
end pad
