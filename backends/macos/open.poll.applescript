-- Polled by the dispatcher after Launch Services opens the file.
-- Returns the active presentation name once PowerPoint reports >=1 open file.
tell application "Microsoft PowerPoint"
    if (count of presentations) > 0 then
        return name of active presentation
    end if
    return ""
end tell
