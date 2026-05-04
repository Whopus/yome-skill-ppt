-- ppt picture.add <slide_index> --path=<image> [--left --top --width --height]
set slideIdx to ({{slide|json}} as integer)
set imgPath to {{path|posix}}
set leftPos to ({{left|json}} as real)
set topPos to ({{top|json}} as real)
set widthArg to {{width|json}}
set heightArg to {{height|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            -- `make new picture` consumes `file name` as POSIX path.
            -- width/height -1 tells PowerPoint to keep native size.
            set useW to -1.0
            set useH to -1.0
            if widthArg is not "" then set useW to (widthArg as real)
            if heightArg is not "" then set useH to (heightArg as real)
            set newShape to make new picture with properties {file name:imgPath, left position:leftPos, top:topPos, width:useW, height:useH}
            return (count of shapes) as string
        end tell
    end tell
end tell
