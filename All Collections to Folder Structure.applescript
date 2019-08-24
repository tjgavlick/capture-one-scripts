-- parent directory to build our nested structure inside
set basePath to POSIX path of (path to home folder) & "/Desktop/tmp/test/"

-- recipe name to enable for exporting
set outputRecipe to "Dropbox Sharing"

-- for testing: max iterations
global maxIterations
set maxIterations to 10
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
on processNestedCollection(thisCollection, parentPath)
	tell application "Capture One 12"
		
		-- ignore tmp and utility directories (my convention is to prefix them with _)
		if name of thisCollection does not start with "_" then
			
			-- standard (non-smart) album in user collections
			if kind of thisCollection is equal to album then
				
				-- create a folder for this album if it doesn't exist, then get a reference
				set thisPath to (parentPath & (name of thisCollection) & "/")
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
						do shell script "rm -f \"" & escapedPath & (name of thisVariant) & "\".*"
						
						process thisVariant
						set currentIteration to currentIteration + 1
						if currentIteration is greater than maxIterations then
							error number -128 -- "user cancelled"
						end if
					end if
				end repeat
			end if
			
			-- nested group in user collections - kick back into this recursive function
			if kind of thisCollection is equal to group and not name of thisCollection is equal to "Recent Imports" then
				repeat with subCollection in (get collections of thisCollection)
					my processNestedCollection(subCollection, parentPath & (name of thisCollection) & "/")
				end repeat
			end if
		end if
		
		
	end tell
end processNestedCollection


tell application "Capture One 12"
	-- make sure only our preferred recipe is active
	repeat with thisRecipe in (get recipes of current document)
		if (name of thisRecipe is outputRecipe) then
			set enabled of thisRecipe to true
		else
			set enabled of thisRecipe to false
		end if
	end repeat
	
	-- get all collections at root level and kick them into recursive listing
	repeat with collectionItem in (get collections of current document)
		my processNestedCollection(collectionItem, basePath)
	end repeat
end tell
