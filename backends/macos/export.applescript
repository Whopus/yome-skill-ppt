-- ppt export --format=pdf|png|jpg --path=<path> [--force]
set fmt to {{format|json}}
set outPath to POSIX file {{path|json}} as string

tell application "Microsoft PowerPoint"
    tell active presentation
        if fmt is "pdf" then
            save in outPath as save as PDF
        else if fmt is "png" or fmt is "images" then
            save in outPath as save as PNG
        else if fmt is "jpg" or fmt is "jpeg" then
            save in outPath as save as JPEG
        else
            error "unsupported format: " & fmt & " (use pdf|png|jpg)" number 7001
        end if
    end tell
    return name of active presentation
end tell
