-- ppt slides
-- Lists all slides in the active presentation as TSV:
-- index<tab>title<tab>shapeCount<tab>layout
tell application "Microsoft PowerPoint"
    tell active presentation
        set slideList to {}
        set sc to count of slides
        repeat with i from 1 to sc
            set s to slide i
            set shapeCount to count of shapes of s
            set titleText to ""
            if shapeCount > 0 then
                try
                    set titleText to content of text range of text frame of shape 1 of s
                end try
            end if
            set ly to layout of s as string
            set end of slideList to (i as string) & tab & titleText & tab & (shapeCount as string) & tab & ly
        end repeat
        set AppleScript's text item delimiters to linefeed
        return slideList as string
    end tell
end tell
