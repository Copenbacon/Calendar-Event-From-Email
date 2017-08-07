set final_dep to current date
set final_check to current date
tell application "Mail"
	set unreadMessages to (get every message of mailbox "EscapiaNet to Calendar Copies" of account "guest@swank-properties.com" whose read status is false)
	repeat with eachMessage in unreadMessages
		set msgcontent to get content of eachMessage as string
		-- set adults to extractBetween(msgcontent, "Adults: ", "Children: ")
		-- set children to extractBetween(msgcontent, "Children: ", "Pets: ")
		-- set pets to extractBetween(msgcontent, "Pets: ", "Guest")
		-- end tell
		tell extractBetween
			set guest_name to extractBetween(msgcontent, "Contact:", "Email:")
			set property_name to extractBetween(msgcontent, "Property", "First")
			set first_night to extractBetween(msgcontent, "Night", "Last")
			set departure_date to extractBetween(msgcontent, "Night", "Adults")
			set adults to extractBetween(msgcontent, "Adults", "Children")
			set children to extractBetween(msgcontent, "Children", "Pets")
			set pets to extractBetween(msgcontent, "Pets", "Guest")
			if property_name contains "Sky View" then
				set property_name to "12P"
			end if
			if property_name contains "Grand View" then
				set property_name to "14P"
			end if
			if property_name contains "Sawyer" then
				set property_name to "58S1"
			end if
			if property_name contains "Shawmut" then
				set property_name to "76L1"
			end if
			if property_name contains "Penthouse" then
				set property_name to "230B4"
			end if
			if property_name contains "Mayhew" then
				set property_name to "230B3"
			end if
			if property_name contains "Roseclair" then
				set property_name to "230B2"
			end if
			if property_name contains "Dorset" then
				set property_name to "230B1"
			end if
			if property_name contains "Stratford" then
				set property_name to "31M1-1"
			end if
			if property_name contains "Monument" then
				set property_name to "4C"
			end if
			if property_name contains "Alameda" then
				set property_name to "25A2"
			end if
			if property_name contains "Bay State" then
				set property_name to "25BSR"
			end if
			if pets contains "None" then
				set pets to "0"
			end if
			set tid to AppleScript's text item delimiters -- save them for later.
			set AppleScript's text item delimiters to {"Last Night: "}
			set placeholder_date to text items of departure_date
			set check_in_date to (text item 1 of departure_date) as string
			set departure_date to (text item 2 of departure_date) as string
			set AppleScript's text item delimiters to {" "}
			set fin_check_in_date to (text item 1 of check_in_date)
			set fin_departure_date to (text item 2 of departure_date)
			set AppleScript's text item delimiters to {"/"} -- reset
			set depart_month to text item 1 of departure_date
			set depart_day to ((text item 2 of departure_date) + 1)
			set month of final_dep to text item 1 of departure_date
			set day of final_dep to (text item 2 of departure_date) + 1
			set year of final_dep to text item 3 of departure_date
			set month of final_check to text item 1 of check_in_date
			set day of final_check to text item 2 of check_in_date
			set year of final_check to text item 3 of check_in_date
			set time of final_check to 17 * hours
			set time of final_dep to 10 * hours
			set AppleScript's text item delimiters to tid
		end tell
		set AppleScript's text item delimiters to tid -- reset
		tell application "Calendar"
			tell calendar "Test"
				make new event with properties {summary:property_name & " " & guest_name & depart_month & "/" & depart_day, start date:final_check, end date:final_check + 1 * hours, location:adults & "-" & children & "-" & pets & " (CPU)"} --reservation check in
				make new event with properties {summary:property_name & " " & guest_name & depart_month & "/" & depart_day, start date:final_dep, end date:final_dep + 1 * hours, location:" Check Out Cleaning (CPU)"} --reservation check out
				make new event with properties {summary:property_name & " " & guest_name & depart_month & "/" & depart_day, start date:final_dep + 7 * days, end date:final_dep + 7 * days + 1 * hours, location:"SD Refund (CPU)"} --SD Refund
			end tell
		end tell
		set read status of eachMessage to true
	end repeat
end tell


-- The following function adapted from the bright minds at https://discussions.apple.com/thread/5188636?tstart=0
on extractBetween(SearchText, startText, endText)
	set tid to AppleScript's text item delimiters -- save them for later.
	set AppleScript's text item delimiters to {" ", return & linefeed, return, linefeed, character id 8233, character id 8232} -- find the first one.
	set liste to every text item of SearchText
	set counter to false
	set extraction to {}
	repeat with subtext in liste
		if subtext contains endText then
			set counter to false
		end if
		if counter is true then
			copy subtext as string to end of extraction
			copy " " as string to end of extraction
		end if
		if subtext contains startText then
			set counter to true
		end if
	end repeat
	set AppleScript's text item delimiters to tid -- back to original values.
	return extraction as string
end extractBetween

