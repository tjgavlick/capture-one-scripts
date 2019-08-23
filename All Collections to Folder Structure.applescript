-- root directory to build our nested structure inside
set basePath to "~/Desktop/tmp/test/"

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

    -- simple album in user collections
    if kind of thisCollection is equal to album then

      -- create a folder for this album if it doesn't exist
      tell application "Finder"
        set thisPath to (parentPath & (name of thisCollection) & "/")
        set escapedPath to my replace(thisPath, " ", "\\ ")
        do shell script "mkdir -p " & escapedPath
        set dir to POSIX file thisPath
      end tell
    end if

    -- nested group in user collections - kick back into this recursive function
    if kind of thisCollection is equal to group and not name of thisCollection is equal to "Recent Imports" then
      repeat with subCollection in (get collections of thisCollection)
        my processNestedCollection(subCollection, parentPath & (name of thisCollection) & "/")
      end repeat
    end if
  end tell
end processNestedCollection


-- get all collections at root level and kick them into recursive listing
tell application "Capture One 12"
  repeat with collectionItem in (get collections of current document)
    my processNestedCollection(collectionItem, basePath)
  end repeat
end tell
