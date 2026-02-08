-- RemindersToCSV.applescript
-- Exports reminders from the macOS Reminders app to a CSV file on the Desktop.
--
-- Usage:
--   From Apple Shortcuts: Pass a list name or "All Lists" as input
--   From Script Editor:   Run directly and a chooser dialog will appear
--
-- Performance: Uses bulk property fetching to minimize Apple Event IPC calls.

-- ============================================================
-- Helper: Escape a single field value for CSV.
-- ============================================================
on escapeCSVField(fieldValue)
	if fieldValue is missing value then return ""
	set fieldText to fieldValue as text
	set needsQuoting to false

	if fieldText contains "\"" then
		set needsQuoting to true
		set AppleScript's text item delimiters to "\""
		set parts to text items of fieldText
		set AppleScript's text item delimiters to "\"\""
		set fieldText to parts as text
		set AppleScript's text item delimiters to ""
	end if

	if fieldText contains "," or fieldText contains linefeed or fieldText contains return then
		set needsQuoting to true
	end if

	if needsQuoting then
		return "\"" & fieldText & "\""
	else
		return fieldText
	end if
end escapeCSVField

-- ============================================================
-- Helper: Format an AppleScript date as YYYY-MM-DD HH:MM
-- ============================================================
on formatDate(theDate)
	if theDate is missing value then return ""

	set y to year of theDate as integer
	set m to month of theDate as integer
	set d to day of theDate as integer
	set h to hours of theDate as integer
	set mins to minutes of theDate as integer

	set yStr to y as text
	if m < 10 then
		set mStr to "0" & (m as text)
	else
		set mStr to m as text
	end if
	if d < 10 then
		set dStr to "0" & (d as text)
	else
		set dStr to d as text
	end if
	if h < 10 then
		set hStr to "0" & (h as text)
	else
		set hStr to h as text
	end if
	if mins < 10 then
		set minStr to "0" & (mins as text)
	else
		set minStr to mins as text
	end if

	return yStr & "-" & mStr & "-" & dStr & " " & hStr & ":" & minStr
end formatDate

-- ============================================================
-- Helper: Convert priority integer to human-readable label.
-- ============================================================
on priorityLabel(priorityValue)
	if priorityValue is missing value or priorityValue = 0 then
		return "None"
	else if priorityValue > 0 and priorityValue < 5 then
		return "High"
	else if priorityValue = 5 then
		return "Medium"
	else
		return "Low"
	end if
end priorityLabel

-- ============================================================
-- Helper: Safe item access
-- ============================================================
on safeItem(idx, theList)
	try
		set val to item idx of theList
		return val
	on error
		return missing value
	end try
end safeItem

-- ============================================================
-- Helper: Build a CSV row from a list of field values
-- ============================================================
on buildCSVRow(fieldList)
	set csvLine to ""
	repeat with i from 1 to count of fieldList
		if i > 1 then set csvLine to csvLine & ","
		set csvLine to csvLine & my escapeCSVField(item i of fieldList)
	end repeat
	return csvLine
end buildCSVRow

-- ============================================================
-- Helper: Export a single reminder list, appending rows to csvRows
-- Returns the number of reminders exported from this list
-- ============================================================
on exportList(aList, csvRows)
	tell application "Reminders"
		set listName to name of aList
		set reminderCount to count of reminders of aList
	end tell

	log "[LIST] " & listName & " - " & reminderCount & " reminders"

	if reminderCount = 0 then return 0

	tell application "Reminders"
		log "  Bulk fetching properties..."
		set allNames to name of every reminder of aList

		try
			set allBodies to body of every reminder of aList
		on error
			set allBodies to {}
		end try

		try
			set allDueDates to due date of every reminder of aList
		on error
			set allDueDates to {}
		end try

		try
			set allCreationDates to creation date of every reminder of aList
		on error
			set allCreationDates to {}
		end try

		try
			set allCompletionDates to completion date of every reminder of aList
		on error
			set allCompletionDates to {}
		end try

		try
			set allCompleted to completed of every reminder of aList
		on error
			set allCompleted to {}
		end try

		try
			set allPriorities to priority of every reminder of aList
		on error
			set allPriorities to {}
		end try

		try
			set allFlagged to flagged of every reminder of aList
		on error
			set allFlagged to {}
		end try
	end tell

	log "  All properties fetched. Building rows..."
	repeat with i from 1 to reminderCount
		set reminderName to item i of allNames

		set reminderBody to my safeItem(i, allBodies)
		if reminderBody is missing value then set reminderBody to ""

		set dueDateVal to my safeItem(i, allDueDates)
		set creationDateVal to my safeItem(i, allCreationDates)
		set completionDateVal to my safeItem(i, allCompletionDates)

		set completedVal to my safeItem(i, allCompleted)
		if completedVal is true then
			set completedStr to "Yes"
		else
			set completedStr to "No"
		end if

		set priorityVal to my safeItem(i, allPriorities)

		set flagVal to my safeItem(i, allFlagged)
		if flagVal is true then
			set flaggedStr to "Yes"
		else
			set flaggedStr to "No"
		end if

		set rowFields to {listName, reminderName, reminderBody, my formatDate(dueDateVal), my formatDate(creationDateVal), my formatDate(completionDateVal), completedStr, my priorityLabel(priorityVal), flaggedStr}
		set end of csvRows to my buildCSVRow(rowFields)

		if i mod 25 = 0 or i = reminderCount then
			log "  Processed " & i & "/" & reminderCount
		end if
	end repeat

	log "  Done with " & listName
	return reminderCount
end exportList

-- ============================================================
-- Main entry point - works from Shortcuts, Script Editor, and osascript
-- ============================================================
on run argv

	log "=============================="
	log "RemindersToCSV Export Starting"
	log "=============================="

	-- Determine which list to export and whether we are interactive
	set listFilter to "All Lists"
	set isInteractive to true

	-- Check if input was passed (Shortcuts or command line)
	try
		if (count of argv) > 0 then
			set listFilter to (item 1 of argv) as text
			set isInteractive to false
			log "[INFO] Received input: " & listFilter
		else
			log "[INFO] No input received, defaulting to All Lists"
		end if
	on error
		log "[INFO] No input received, defaulting to All Lists"
	end try

	log "[INFO] Export target: " & listFilter

	-- Build the timestamped output filename
	set todayDate to current date
	set dateStamp to my formatDate(todayDate)
	-- Make filename-safe
	set AppleScript's text item delimiters to " "
	set dateStamp to text items of dateStamp
	set AppleScript's text item delimiters to "_"
	set dateStamp to dateStamp as text
	set AppleScript's text item delimiters to ":"
	set dateStamp to text items of dateStamp
	set AppleScript's text item delimiters to ""
	set dateStamp to dateStamp as text

	if listFilter = "All Lists" then
		set exportFileName to "Reminders_Export_" & dateStamp & ".csv"
	else
		-- Include list name in filename (strip unsafe chars)
		set safeName to listFilter
		set AppleScript's text item delimiters to "/"
		set safeName to text items of safeName
		set AppleScript's text item delimiters to "_"
		set safeName to safeName as text
		set AppleScript's text item delimiters to ":"
		set safeName to text items of safeName
		set AppleScript's text item delimiters to "_"
		set safeName to safeName as text
		set AppleScript's text item delimiters to ""
		set exportFileName to "Reminders_" & safeName & "_" & dateStamp & ".csv"
	end if

	set desktopPath to (path to desktop folder as text)
	set exportFilePath to desktopPath & exportFileName
	set posixPath to POSIX path of exportFilePath
	log "[INFO] Output file: " & posixPath

	-- CSV header row
	set csvHeader to "List Name,Title,Notes,Due Date,Creation Date,Completion Date,Completed,Priority,Flagged"
	set csvRows to {csvHeader}
	set totalCount to 0

	try
		tell application "Reminders"
			log "[INFO] Connecting to Reminders app..."

			if listFilter = "All Lists" then
				set listsToExport to every list
			else
				try
					set listsToExport to {list listFilter}
				on error
					log "[ERROR] List not found: " & listFilter
					if isInteractive then
						display dialog "List not found: " & listFilter buttons {"OK"} default button "OK" with icon stop
					end if
					return "Error: List not found"
				end try
			end if

			set listCount to count of listsToExport
			log "[INFO] Exporting " & listCount & " list(s)"
		end tell

		-- Export each list
		repeat with aList in listsToExport
			set exported to my exportList(aList, csvRows)
			set totalCount to totalCount + exported
		end repeat

		log "[INFO] All lists processed. Total reminders: " & totalCount

		if totalCount = 0 then
			log "[WARN] No reminders found."
			if isInteractive then
				display dialog "No reminders found in the selected list(s)." buttons {"OK"} default button "OK" with icon note
			end if
			return "No reminders found"
		end if

		log "[INFO] Building CSV content..."
		set AppleScript's text item delimiters to linefeed
		set csvContent to csvRows as text
		set AppleScript's text item delimiters to ""

		-- Write the CSV file
		log "[INFO] Writing CSV file to disk..."
		try
			do shell script "printf '%s' " & quoted form of csvContent & " > " & quoted form of posixPath
			log "[INFO] File written successfully."
		on error writeErr
			log "[ERROR] Failed to write file: " & writeErr
			if isInteractive then
				display dialog "Error writing file: " & writeErr buttons {"OK"} default button "OK" with icon stop
			end if
			return "Error: " & writeErr
		end try

		-- Show completion dialog (only when running interactively)
		if isInteractive then
			set msg to "Export complete!" & return & return
			set msg to msg & "List: " & listFilter & return
			set msg to msg & "Reminders exported: " & totalCount & return
			set msg to msg & "File saved to:" & return & posixPath
			display dialog msg buttons {"OK"} default button "OK" with icon note
		end if

		-- Always open the destination folder in Finder
		tell application "Finder"
			open folder (path to desktop folder)
			activate
		end tell

		log "[INFO] Export complete. " & totalCount & " reminders exported."
		log "=============================="
		log "RemindersToCSV Export Finished"
		log "=============================="

		-- Return the file path to Shortcuts for further actions
		return posixPath

	on error errMsg number errNum
		log "[ERROR] Export failed - Error " & errNum & ": " & errMsg
		if isInteractive then
			set errDisplay to "An error occurred during export:" & return & return
			set errDisplay to errDisplay & "Error " & errNum & ": " & errMsg
			display dialog errDisplay buttons {"OK"} default button "OK" with icon stop
		end if
		return "Error: " & errMsg
	end try
end run
