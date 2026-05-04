-- ppt shape.resize <slide_index> --shape=<idx> [--width=W] [--height=H]
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)
set widthArg to {{width|json}}
set heightArg to {{height|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set sh to shape shapeIdx
            if widthArg is not "" then
                set width of sh to (widthArg as real)
            end if
            if heightArg is not "" then
                set height of sh to (heightArg as real)
            end if
            return "resized"
        end tell
    end tell
end tell
