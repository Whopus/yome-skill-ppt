-- ppt table.add <slide_index> --rows=R --cols=C [--left --top --width --height]
set slideIdx to ({{slide|json}} as integer)
set rowsArg to ({{rows|json}} as integer)
set colsArg to ({{cols|json}} as integer)
set leftPos to ({{left|json}} as real)
set topPos to ({{top|json}} as real)
set widthVal to ({{width|json}} as real)
set heightVal to ({{height|json}} as real)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set newShape to make new shape table with properties {number of rows:rowsArg, number of columns:colsArg, left position:leftPos, top:topPos, width:widthVal, height:heightVal}
            return (count of shapes) as string
        end tell
    end tell
end tell
