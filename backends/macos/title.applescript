-- ppt title <slide_index> --text=<text>
-- Sets the text of shape 1 on the slide (the title placeholder by convention).
set slideIdx to ({{slide|json}} as integer)
set newText to {{text|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        set content of text range of text frame of shape 1 of slide slideIdx to newText
    end tell
end tell
return "updated"
