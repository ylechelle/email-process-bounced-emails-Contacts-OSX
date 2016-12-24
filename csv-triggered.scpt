
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

tell application "Mail" to set theMessages to get selection

set {failedAddresses, failedDates} to my mailTest(theMessages, deleteSelectedMessages)
set {cCount, attnCount} to {0, 0}
if failedAddresses ≠ {} then set {cCount, attnCount} to my cleanAddressBook(failedAddresses, failedDates)

my logIt("Deprecated email addresses found: " & (count failedAddresses) & return & "Cleaned contacts: " & cCount & return & "Contacts added to Need Attention: " & attnCount, currentLog, "run")

do shell script "open " & quoted form of currentLog
--end perform mail action with messages
--end using terms from

-------------------------------------------------------------------------------------
-- Handler to identify and parse deprecated emails
-- If removeAddresses does not equal missing value, process mailboxes
on mailTest(myMessages, removeSelection)
	
	script o
		property oMessages : myMessages
		property targetEmails : {}
		property targetDates : {}
	end script
	
	
	local senderAddress, dateReceived, mID, mMbox, mAccount, mContent
	repeat with aMessage in o's oMessages
		try
			
			set aMessage to (contents of aMessage)
			
			tell application "Mail" to set {senderAddress, dateReceived, mID, mMbox, mAccount, mContent} to {extract address from sender, date received, message id, mailbox's name, mailbox's account's name, content} of aMessage
			
			-- 20141226 : identify single quotes ''
			
			-- I believe starts with "MAILER-DAEMON@" will catch all instances.
			-- NEW dec 2015: when dealing with a batch of failed delivery addresses in a single mailer-daemon email, replace the "AND" after 'with "postmaster@"' below with "OR" to catch them all after truncating the email eventually in a new draft and run the script on this draft
			-- do NOT forget to switch back to "AND" afterwards
			
			if (senderAddress starts with "MAILER-DAEMON@" and mContent does not contain "user does not exist" and mContent does not contain "mailbox unavailable") or (senderAddress starts with "postmaster@" and (mContent contains "550" or mContent contains "5.1.1")) then
				--set tgtEmails to paragraphs of (do shell script "osascript -e 'tell app \"Mail\" to get content of first message of mailbox \"" & mMbox & "\" of account \"" & mAccount & "\" whose message id = \"" & mID & "\"' | sed -En '/^<[^<][^@]+@[^>]+>/ { s/<|>.*//g ; p ; }'")
				--Pick out lines beginning with a single "<" and containing both "@" and ">" in that order and return those lines with "<" and (">" and anything after it) removed.
				
				--set tgtEmails to paragraphs of (do shell script "osascript -e 'tell app \"Mail\" to get content of first message of mailbox \"" & mMbox & "\" of account \"" & mAccount & "\" whose message id = \"" & mID & "\"' | sed -En 's/^<?([^<>: ]+@[^<>: ]+).*/\\1/p'")
				--Assume that in these messages, any group of non-space characters surrounding a "@" is likely to be an e-mail address, and if we also assume that each address we want comes at the beginning of a line and is immediately followed by an angle bracket, a colon, a space, or the end of the line.
				
				-- set tgtEmails to (do shell script "osascript -e 'tell app \"Mail\" to get content of first message of mailbox \"" & mMbox & "\" of account \"" & mAccount & "\" whose message id = \"" & mID & "\"' | sed -En 's/^<?((\\([^\\)]*\\))?[^<>: ]+(\\([^\\)]*\\))?@(\\([^\\)]*\\))?([[][^]]+[]]|[[:alnum:].-]+[[:alpha:]])(\\([^\\)]*\\))?).*/\\1/p'")
				
				--Make the angle brackets round the e-mail addresses optional while at the same time identifying the ends of the addresses and actively recognising as many of the unlikely but allowed address forms (<http://en.wikipedia.org/wiki/E-mail_address#Syntax>) as possible.
				
				-- Excludes any address which contains "Dewost"
				-- set tgtEmails to (do shell script "osascript -e 'tell app \"Mail\" to get content of first message of mailbox \"" & mMbox & "\" of account \"" & mAccount & "\" whose message id = \"" & mID & "\"' | sed -En '/[Dd]ewost/ !s/^<?((\\([^\\)]*\\))?[^<>: ]+(\\([^\\)]*\\))?@(\\([^\\)]*\\))?([[][^]]+[]]|[[:alnum:].-]+[[:alpha:]])(\\([^\\)]*\\))?).*/\\1/p'")
				
				-- Excludes any address which contains "Dewost". Otherwise catches addresses at the beginnings of lines or indented with white space.
				
				set tgtEmails to (do shell script "osascript -e 'tell app \"Mail\" to get content of first message of mailbox \"" & mMbox & "\" of account \"" & mAccount & "\" whose message id = \"" & mID & "\"' | sed -En '/[Dd]ewost/ !s/^[[:blank:]]*<?((\\([^\\)]*\\))?[^<>: ]+(\\([^\\)]*\\))?@(\\([^\\)]*\\))?([[][^]]+[]]|[[:alnum:].-]+[[:alpha:]])(\\([^\\)]*\\))?).*/\\1/p'")
				
				-- postmaster@ & -550 or 5.1.1 5.2.2
				--dialog stats
				
				-- Remove leading whitespace
				set tgtEmails to paragraphs of (do shell script "sed 's/^[^[:alnum:]]*//' <<< " & quoted form of tgtEmails)
				
				set o's targetEmails to o's targetEmails & tgtEmails
				repeat (count tgtEmails) times
					set o's targetDates to o's targetDates & (dateReceived as text)
				end repeat
				if removeSelection then delete aMessage
			end if
			
		on error errMsg number errNum
			my logIt("mailTest Handler: " & errMsg & return & "Error number" & errNum, currentLog, "run")
		end try
	end repeat
	
	my logIt(o's targetEmails, currentLog, "run")
	return {o's targetEmails, o's targetDates}
	beep
end mailTest

-------------------------------------------------------------------------------------

-- Handler to clean up Address Book
on cleanAddressBook(deprecatedAddresses, bounceDates)
	script p
		property pAddresses : deprecatedAddresses
		property pDates : bounceDates
	end script
	
	set cleanedCount to 0
	set attentionCount to 0
	
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
			set anAddress to (item i of p's pAddresses)
			
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
						set myContact's note to "deprecated email address: " & anAddress & " bounced on: " & (item i of p's pDates) & return & myContact's note
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
		end repeat
	end tell
	
	return {cleanedCount, attentionCount}
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

