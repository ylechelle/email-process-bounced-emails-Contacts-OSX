
-- Delete all messages from selection containing deprecated addresses
property deleteSelectedMessages : true
-- Resend message to other email addresses (use only whith a mail rule, avoid when running script on a batch of messages)
property doSendAgain : false

property myLabel : "deprecated"
property groupA : "Need Attention"
property groupB : "deprecated email"
-------------------------------------------------------------------------------------

--using terms from application "Mail"
--on perform mail action with messages theMessages for rule theRule
beep
global currentLog
set currentLog to my makeLog()

--load csv file here
set theFile to (choose file with prompt "Select a MailChimp bounced emails file (csv):")
open for access theFile
set fileContents to (read theFile)
close access theFile

--parse email addresses into failedAddresses array
set failedAddresses to my csvToList(fileContents, {separator:","}, {trimming:true})
set failedDates to {}

--
set {cCount, attnCount} to {0, 0}
if failedAddresses ≠ {} then set {pCount, cCount, attnCount} to my cleanAddressBook(failedAddresses, failedDates)

my logIt("Deprecated email addresses found: " & pCount & return & "Cleaned contacts: " & cCount & return & "Contacts added to Need Attention: " & attnCount, currentLog, "run")

do shell script "open " & quoted form of currentLog
--end perform mail action with messages
--end using terms from

-- Handler to clean up Address Book
on cleanAddressBook(deprecatedAddresses)
	script p
		property pAddresses : deprecatedAddresses
	end script
	
	set processCount to 0
	set cleanedCount to 0
	set attentionCount to 0
	set today to date string of (current date)
	
	-- Create groups in Address Book
	tell application "Contacts"
		activate
		repeat with myGroup in {groupA, groupB}
			set myGroup to contents of myGroup
			if not (exists (every group whose name = myGroup)) then
				make new group with properties {name:myGroup}
				save
			end if
		end repeat
		
		repeat with i from 1 to (count of p's pAddresses)
			set aRow to (item i of p's pAddresses)
			set anAddress to (item 1 of aRow)
			
			if "@" is in anAddress then
				set processCount to processCount + 1

				if exists (first person whose value of emails contains anAddress) then
					try
						set myContact to (first person whose value of emails contains anAddress)
						
						--only resend if client has another address
						if (myContact's emails count) > 1 and doSendAgain then
							set replaceAddress to (first email of myContact whose value ≠ anAddress)'s value
							tell me to sendAgain(anAddress, replaceAddress)
							delay 3
						end if
						
						set contactName to myContact's name
						set contactEmail to (first email of myContact whose value = anAddress)
						set contactEmail's label to myLabel
						
						set emailCount to (count of myContact's emails)
						if emailCount = 1 then
							set groupX to (first group whose name = groupA)
						else
							set groupX to (first group whose name = groupB)
						end if
						
						set removeContact to true
						if removeContact then
							delete contactEmail
							if (myContact's id) is not in (people's id of groupX) then
								add myContact to groupX
								save
							end if
							if myContact's note = missing value then set myContact's note to ""
							set myContact's note to "deprecated email address: " & anAddress & " bounced on: " & today & return & myContact's note
							save
							
							if emailCount = 1 then
								set attentionCount to attentionCount + 1
							else
								set cleanedCount to cleanedCount + 1
							end if
						end if
						
					on error errMsg number errNum
						if errNum = -128 then
							error number -128
						else
							my logIt("cleanAddressBook Handler: " & errMsg & return & "Error number" & errNum, currentLog, "run")
						end if
					end try
				end if
			end if
		end repeat
	end tell
	
	return {processCount, cleanedCount, attentionCount}
end cleanAddressBook
-------------------------------------------------------------------------------------

on sendAgain(findAddress, replaceAddress)
	try
		tell application "Mail"
			activate
			set resendMessages to (sent mailbox's messages whose date sent > ((current date) - 1 * days) and recipients's address contains findAddress)
		end tell
		
		if the resendMessages ≠ {} then
			tell application "Mail"
				open first item of the resendMessages
				activate
			end tell
			
			tell application "System Events" to tell process "Mail"
				keystroke "d" using {command down, shift down}
				try
					set value of text field 1 of scroll area 3 of front window to replaceAddress
				end try
				try
					set value of text field 1 of scroll area 2 of front window to ""
				end try
				try
					set value of text field 1 of scroll area 4 of front window to ""
				end try
				try
					set value of text field 1 of scroll area 1 of front window to ""
				end try
			end tell
		end if
		
	on error errMsg number errNum
		my logIt("sendAgain Handler: " & errMsg & return & "Error number" & errNum, currentLog, "run")
	end try
end sendAgain
-------------------------------------------------------------------------------------
on makeLog()
	set logFolder to POSIX path of (path to desktop as text) & "Mail Log"
	do shell script "mkdir -p " & quoted form of logFolder
	
	set cd to {year, month, day, time string} of (current date)
	set cd's item 2 to text -2 thru -1 of ("0" & (cd's item 2 as number))
	set cd to do shell script "sed 's/[:APM]//g' <<< " & quoted form of (cd as text)
	
	set logPath to logFolder & "/" & cd
	
	return logPath
end makeLog

on logIt(theMessage, logPath, action)
	if class of theMessage = list then
		set AppleScript's text item delimiters to ", "
		set theMessage to theMessage as text
		set AppleScript's text item delimiters to {""}
	end if
	do shell script "echo " & quoted form of theMessage & " >> " & quoted form of logPath
	if action = "quit" then error number -128
end logIt

---------- csvToList function by Nigel Garvey - http://macscripter.net/viewtopic.php?pid=125444#p125444

on csvToList(csvText, implementation)
   -- The 'implementation' parameter must be a record. Leave it empty ({}) for the default assumptions: ie. comma separator, leading and trailing spaces in unquoted fields not to be trimmed. Otherwise it can have a 'separator' property with a text value (eg. {separator:tab}) and/or a 'trimming' property with a boolean value ({trimming:true}).
   set {separator:separator, trimming:trimming} to (implementation & {separator:",", trimming:false})
   
   script o -- Lists for fast access.
       property qdti : getTextItems(csvText, "\"")
       property currentRecord : {}
       property possibleFields : missing value
       property recordList : {}
   end script
   
   -- o's qdti is a list of the CSV's text items, as delimited by double-quotes.
   -- Assuming the convention mentioned above, the number of items is always odd.
   -- Even-numbered items (if any) are quoted field values and don't need parsing.
   -- Odd-numbered items are everything else. Empty strings in odd-numbered slots
   -- (except at the beginning and end) indicate escaped quotes in quoted fields.
   
   set astid to AppleScript's text item delimiters
   set qdtiCount to (count o's qdti)
   set quoteInProgress to false
   considering case
       repeat with i from 1 to qdtiCount by 2 -- Parse odd-numbered items only.
           set thisBit to item i of o's qdti
           if ((count thisBit) > 0) or (i is qdtiCount) then
               -- This is either a non-empty string or the last item in the list, so it doesn't
               -- represent a quoted quote. Check if we've just been dealing with any.
               if (quoteInProgress) then
                   -- All the parts of a quoted field containing quoted quotes have now been
                   -- passed over. Coerce them together using a quote delimiter.
                   set AppleScript's text item delimiters to "\""
                   set thisField to (items a thru (i - 1) of o's qdti) as string
                   -- Replace the reconstituted quoted quotes with literal quotes.
                   set AppleScript's text item delimiters to "\"\""
                   set thisField to thisField's text items
                   set AppleScript's text item delimiters to "\""
                   -- Store the field in the "current record" list and cancel the "quote in progress" flag.
                   set end of o's currentRecord to thisField as string
                   set quoteInProgress to false
               else if (i > 1) then
                   -- The preceding, even-numbered item is a complete quoted field. Store it.
                   set end of o's currentRecord to item (i - 1) of o's qdti
               end if
               
               -- Now parse this item's field-separator-delimited text items, which are either non-quoted fields or stumps from the removal of quoted fields. Any that contain line breaks must be further split to end one record and start another. These could include multiple single-field records without field separators.
               set o's possibleFields to getTextItems(thisBit, separator)
               set possibleFieldCount to (count o's possibleFields)
               repeat with j from 1 to possibleFieldCount
                   set thisField to item j of o's possibleFields
                   if ((count thisField each paragraph) > 1) then
                       -- This "field" contains one or more line endings. Split it at those points.
                       set theseFields to thisField's paragraphs
                       -- With each of these end-of-record fields except the last, complete the field list for the current record and initialise another. Omit the first "field" if it's just the stub from a preceding quoted field.
                       repeat with k from 1 to (count theseFields) - 1
                           set thisField to item k of theseFields
                           if ((k > 1) or (j > 1) or (i is 1) or ((count trim(thisField, true)) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
                           set end of o's recordList to o's currentRecord
                           set o's currentRecord to {}
                       end repeat
                       -- With the last end-of-record "field", just complete the current field list if the field's not the stub from a following quoted field.
                       set thisField to end of theseFields
                       if ((j < possibleFieldCount) or ((count thisField) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
                   else
                       -- This is a "field" not containing a line break. Insert it into the current field list if it's not just a stub from a preceding or following quoted field.
                       if (((j > 1) and ((j < possibleFieldCount) or (i is qdtiCount))) or ((j is 1) and (i is 1)) or ((count trim(thisField, true)) > 0)) then set end of o's currentRecord to trim(thisField, trimming)
                   end if
               end repeat
               
               -- Otherwise, this item IS an empty text representing a quoted quote.
           else if (quoteInProgress) then
               -- It's another quote in a field already identified as having one. Do nothing for now.
           else if (i > 1) then
               -- It's the first quoted quote in a quoted field. Note the index of the
               -- preceding even-numbered item (the first part of the field) and flag "quote in
               -- progress" so that the repeat idles past the remaining part(s) of the field.
               set a to i - 1
               set quoteInProgress to true
           end if
       end repeat
   end considering
   
   -- At the end of the repeat, store any remaining "current record".
   if (o's currentRecord is not {}) then set end of o's recordList to o's currentRecord
   set AppleScript's text item delimiters to astid
   
   return o's recordList
end csvToList

-- Get the possibly more than 4000 text items from a text.
on getTextItems(txt, delim)
   set astid to AppleScript's text item delimiters
   set AppleScript's text item delimiters to delim
   set tiCount to (count txt's text items)
   set textItems to {}
   repeat with i from 1 to tiCount by 4000
       set j to i + 3999
       if (j > tiCount) then set j to tiCount
       set textItems to textItems & text items i thru j of txt
   end repeat
   set AppleScript's text item delimiters to astid
   
   return textItems
end getTextItems

-- Trim any leading or trailing spaces from a string.
on trim(txt, trimming)
   if (trimming) then
       repeat with i from 1 to (count txt) - 1
           if (txt begins with space) then
               set txt to text 2 thru -1 of txt
           else
               exit repeat
           end if
       end repeat
       repeat with i from 1 to (count txt) - 1
           if (txt ends with space) then
               set txt to text 1 thru -2 of txt
           else
               exit repeat
           end if
       end repeat
       if (txt is space) then set txt to ""
   end if
   
   return txt
end trim

