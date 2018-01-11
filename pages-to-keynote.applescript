global filepath, charactersPerPage, currentText, keynoteDocument, currentTextStyles

set filepath to ""
set charactersPerPage to 400
set currentText to ""
set currentTextStyles to {}

property slideWidth : 1920
property slideHeight : 1080
property fontSize : 95

property keynoteTheme : "Black"

on replaceChars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replaceChars

on flushPage()
	log "FLUSHING TO KEYNOTE: " & currentText & ". Count: " & (length of currentText)
	-- create new slide at the end of the document
	tell application "Keynote"
		tell keynoteDocument
			set newSlide to make new slide with properties {base slide:master slide "Blank"} at the end of slides
			tell newSlide
				set slideText to make new text item with properties {object text:currentText}
				set the size of object text of slideText to fontSize
				
				repeat with i from 1 to length of currentText
					set currentTextItemStyle to item i of currentTextStyles
					tell slideText
						set the font of character i of object text to (the font of currentTextItemStyle)
						
						-- since we use a black theme, the text needs to be readible
						if the color of currentTextItemStyle is not {0, 0, 0} then
							set the color of character i of object text to (the color of currentTextItemStyle)
						end if
						
					end tell
				end repeat
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
		
		set keynoteDocument to Â
			make new document with properties Â
				{document theme:theme keynoteTheme, width:1920, height:1080}
		tell keynoteDocument
			set the base slide of the first slide to master slide "Title & Subtitle"
			tell the first slide
				set the object text of the default title item to "Presentation"
				set the object text of the default body item to "It starts on the next slide"
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

display dialog Â
	"Please, choose a Pages file to convert to Keynote. " & Â
	return & Â
	return & Â
	Â
		"If you need to convert a Microsoft Word document, then open it in Pages and save as a pages file." buttons {"Cancel", "Continue"} Â
	default button Â
	"Continue" cancel button "Cancel"


open {choose file}
