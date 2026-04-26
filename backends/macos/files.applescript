-- ppt files
-- Lists all open presentations: TSV "name<tab>slideCount".
tell application "Microsoft PowerPoint"
    set presList to {}
    set presCount to count of presentations
    repeat with i from 1 to presCount
        set p to presentation i
        set pName to name of p
        set sc to count of slides of p
        set end of presList to pName & tab & (sc as string)
    end repeat
    set AppleScript's text item delimiters to linefeed
    return presList as string
end tell
