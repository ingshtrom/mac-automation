global filepath, charactersPerPage, currentText, keynoteDocument, currentTextStyles

set filepath to ""
set charactersPerPage to 175 -- CHAR_PER_LINE * <the_number_of_lines_that_fit_per_slide>
set currentText to ""
set currentTextStyles to {}

property fontSize : 95
property charSpaceCount : 10

property CHAR_PER_LINE : 35
property TEXT_FONT : "Arial Bold"

property keynoteTheme : "Black"

-- copied from somewhere. don't remember
on replaceChars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replaceChars

on addLineBreaksToCurrentText()
	set updatedText to ""
	set updatedStyles to {}
	set shouldBreakNextSpace to false
	
	repeat with i from 1 to length of currentText
		set curChar to character i of currentText
		
		if i mod CHAR_PER_LINE is equal to 0 then
			set shouldBreakNextSpace to true
		end if
		
		if shouldBreakNextSpace is equal to true and curChar is equal to " " then
			set updatedText to (updatedText & curChar & return & return)
			set updatedStyles to updatedStyles & {item i of currentTextStyles} & {item i of currentTextStyles} & {item i of currentTextStyles}
			set shouldBreakNextSpace to false
		else
			set updatedText to (updatedText & curChar)
			set updatedStyles to updatedStyles & {item i of currentTextStyles}
		end if
	end repeat
	
	set currentText to updatedText
	set currentTextStyles to updatedStyles
end addLineBreaksToCurrentText

on flushPage()
	log "FLUSHING TO KEYNOTE: " & currentText & ". Count: " & (length of currentText)
	-- create new slide at the end of the document
	tell application "Keynote"
		tell keynoteDocument
			set newSlide to make new slide with properties {base slide:master slide "Blank"} at the end of slides
			tell newSlide
				my addLineBreaksToCurrentText()
				set slideText to make new text item with properties {object text:currentText}
				set the size of object text of slideText to fontSize
				set the font of object text of slideText to TEXT_FONT
				
				repeat with i from 1 to length of currentText
					set currentTextItemStyle to item i of currentTextStyles
					tell slideText
						-- set the font of character i of object text to (the font of currentTextItemStyle)
						
						-- since we use a black theme, the text needs to be readible
						if the color of currentTextItemStyle is not {0, 0, 0} then
							set the color of character i of object text to (the color of currentTextItemStyle)
						end if
						
					end tell
				end repeat
				tell application "System Events"
					repeat with i from 1 to 10
						keystroke "]" using {command down, option down}
					end repeat
				end tell
			end tell
		end tell
	end tell
	-- delete any current default items in the slide
	-- add currentText to the slide
	set currentText to ""
	set currentTextStyles to {}
end flushPage


on createNewKeynoteDocument()
	tell application "Keynote"
		activate
		
		-- GET THE THEME NAMES
		set the themeNames to the name of every theme
		
		log "Available Keynote themes: " & themeNames
		log "My theme: " & keynoteTheme
		
		set keynoteDocument to make new document with properties {document theme:theme keynoteTheme, width:1920, height:1080}
		tell keynoteDocument
			set the base slide of the first slide to master slide "Title & Subtitle"
			tell the first slide
				set the object text of the default title item to "Generated on " & (current date)
			end tell
		end tell
	end tell
end createNewKeynoteDocument


on mainFn(location)
	try
		set filepath to location
		createNewKeynoteDocument()
		
		set thisPOSIXPath to (the POSIX path of filepath)
		log "FILE TO OPEN: " & thisPOSIXPath
		
		do shell script "open '" & thisPOSIXPath & "'"
		
		delay (10)
		
		tell application "Pages"
			tell the front document
				tell the body text
					set charRefs to a reference to every character
					
					
					--					log "" & (the color of the first word of the first paragraph)
					
					--					repeat with paragraphItem in bodyTextByParagraph
					--						log "para " & (the properties of paragraphItem)
					repeat with charRef in charRefs
						set charItem to (contents of charRef)
						
						set potentialNewCharacterCount to (length of currentText) + 1
						
						if potentialNewCharacterCount is greater than or equal to charactersPerPage and charItem is equal to " " then
							my flushPage()
						end if
						
						set currentText to currentText & charItem
						set currentTextStyles to currentTextStyles & {charRef}
					end repeat
					--				end repeat
					
					my flushPage()
				end tell
			end tell
			
			quit
		end tell
		
		tell application "Keynote"
			log "User is being prompted to save the keynote somewhere..."
			save front document
			quit
		end tell
	on error errStr number errorNumber
		quit application "Pages" without saving
		quit application "Keynote" without saving
		error errStr number errorNumber
	end try
end mainFn

on open droppedItems
	repeat with a from 1 to length of droppedItems
		set theCurrentDroppedItem to item a of droppedItems
		log "DROPPED ITEM: " & theCurrentDroppedItem
		my mainFn(theCurrentDroppedItem)
	end repeat
end open

display dialog "Please, choose a Pages file to convert to Keynote. " & return & return & "If you need to convert a Microsoft Word document, then open it in Pages and save as a pages file." buttons {"Cancel", "Continue"} default button "Continue" cancel button "Cancel"


open {choose file}
