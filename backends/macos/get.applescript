-- ppt get <slide_index>
-- Returns TSV: index<tab>name<tab>hasText<tab>content for every shape.
set slideIdx to ({{slide|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set shapeCount to count of shapes
            set lineList to {}
            repeat with i from 1 to shapeCount
                set sh to shape i
                set shName to name of sh
                set hasText to has text frame of sh
                set textContent to ""
                if hasText then
                    try
                        set textContent to content of text range of text frame of sh
                    end try
                end if
                set end of lineList to (i as string) & tab & shName & tab & (hasText as string) & tab & textContent
            end repeat
            set AppleScript's text item delimiters to linefeed
            return lineList as string
        end tell
    end tell
end tell
