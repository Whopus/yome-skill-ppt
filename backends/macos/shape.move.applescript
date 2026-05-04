-- ppt shape.move <slide_index> --shape=<idx> [--left=X] [--top=Y]
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)
set leftArg to {{left|json}}
set topArg to {{top|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set sh to shape shapeIdx
            if leftArg is not "" then
                set left position of sh to (leftArg as real)
            end if
            if topArg is not "" then
                set top of sh to (topArg as real)
            end if
            return "moved"
        end tell
    end tell
end tell
