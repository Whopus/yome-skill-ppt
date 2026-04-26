-- ppt set <slide_index> --shape=<idx> --text=<text>
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)
set newText to {{text|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        set content of text range of text frame of shape shapeIdx of slide slideIdx to newText
    end tell
end tell
return "updated"
