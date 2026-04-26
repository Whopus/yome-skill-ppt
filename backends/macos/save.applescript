-- ppt save [--path=<save_as>] [--force]
-- If --path is empty -> plain save. Otherwise save as that path in OOXML.
set saveAsPath to {{path|json}}
tell application "Microsoft PowerPoint"
    if saveAsPath is "" then
        tell active presentation
            save
        end tell
    else
        set destPath to POSIX file saveAsPath as string
        tell active presentation
            save in destPath as save as Open XML presentation
        end tell
    end if
    return name of active presentation
end tell
