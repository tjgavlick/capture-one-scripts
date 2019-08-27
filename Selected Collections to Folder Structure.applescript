-- parent directory to build our nested structure inside
set basePath to POSIX path of (path to home folder) & "/Dropbox/Photos/"

-- recipe name to enable for exporting
set outputRecipe to "Dropbox Sharing"

-- variable to track identifier of current collection
global selectedCollection

-- for safety and/or testing: max iterations
global maxIterations
set maxIterations to 50000
global currentIteration
set currentIteration to 1


-- string find and replace helper
on replace(theText, theSearchString, theReplacementString)
	set AppleScript's text item delimiters to theSearchString
	set theTextItems to every text item of theText
	set AppleScript's text item delimiters to theReplacementString
	set theText to theTextItems as string
	set AppleScript's text item delimiters to ""
	return theText
end replace


-- handle both albums and sub-collections inside a given collection
on processNestedCollection(thisCollection, parentPath, isWorkingCollection)
	set currentIsWorkingCollection to isWorkingCollection
	
	tell application "Capture One 12"
		
		if id of thisCollection is equal to id of selectedCollection then set currentIsWorkingCollection to true
		
		-- standard (non-smart) album in user collections
		if kind of thisCollection is equal to album then
			
			set thisPath to (parentPath & (name of thisCollection) & "/")
			
			-- process only variants in the selected collection (or children of the selected collection)
			if currentIsWorkingCollection is true then
				
				-- create a folder for this album if it doesn't exist, then get a reference
				set escapedPath to my replace(thisPath, "\"", "\\\"")
				do shell script "mkdir -p \"" & escapedPath & "\""
				set dir to POSIX file thisPath as alias
				
				-- set recipe output to our directory
				set output of current document to dir
				
				-- process collection variants
				repeat with thisVariant in (get variants of thisCollection)
					
					-- only process non-rejects and non-helper variants
					if color tag of thisVariant is equal to 0 or color tag of thisVariant is equal to 4 then
						
						-- Capture One appends numbers to the output filename if there is a conflict, but we want overwriting
						-- so, delete existing file if it already exists
						-- we have to assume the image name is inside the output filename, as C1 does not (currently) give us
						-- a way to get the expected output filename of this variant from the recipe or queue
						do shell script "rm -f \"" & escapedPath & "\"*\"" & (name of thisVariant) & "\"*"
						
						process thisVariant
						set currentIteration to currentIteration + 1
						if currentIteration is greater than maxIterations then
							error number -128 -- "user cancelled"
						end if
					end if
				end repeat
			end if
		end if
		
		-- nested group in user collections - kick back into this recursive function
		if kind of thisCollection is equal to group and not name of thisCollection is equal to "Recent Imports" then
			repeat with subCollection in (get collections of thisCollection)
				my processNestedCollection(subCollection, parentPath & (name of thisCollection) & "/", currentIsWorkingCollection)
			end repeat
		end if
		
		
	end tell
end processNestedCollection


tell application "Capture One 12"
	set selectedCollection to current collection of current document
	
	-- make sure only our preferred recipe is active
	repeat with thisRecipe in (get recipes of current document)
		if (name of thisRecipe is outputRecipe) then
			set enabled of thisRecipe to true
		else
			set enabled of thisRecipe to false
		end if
	end repeat
	
	-- Capture One appends the selected collection to the return of all collections
	-- so, trim it off to avoid duplicate work
	set allCollections to (get collections of current document)
	set allCollections to items 1 through -2 of allCollections
	
	-- set up progress bar for collection tree scan
	set progress total units to the length of allCollections
	set progress completed units to 0
	set progress text to "Crawling collections ..."
	
	-- pause queue to free up C1's resources to traverse our collections and variants
	set processing queue enabled of current document to false
	
	-- kick collections into recursive listing
	repeat with collectionItem in allCollections
		my processNestedCollection(collectionItem, basePath, false)
		set progress completed units to ((progress completed units) + 1)
	end repeat
	
	-- start crunching
	set processing queue enabled of current document to true
	
	display dialog "Added " & (currentIteration - 1) & " variants to process queue"
end tell
