-- ppt shape.add <slide_index> --type=<rectangle|oval|roundedRect|triangle|rightArrow|star5|pentagon|diamond|hexagon|cloud|lightningBolt|heart>
--                               [--left --top --width --height --text]
--
-- The dispatcher substitutes {{type|autoshape}} with the full AppleScript
-- enumeration literal like "autoshape rectangle" based on the user-friendly
-- alias above. Unknown aliases fall through to "autoshape rectangle".
set slideIdx to ({{slide|json}} as integer)
set leftPos to ({{left|json}} as real)
set topPos to ({{top|json}} as real)
set widthVal to ({{width|json}} as real)
set heightVal to ({{height|json}} as real)
set txtArg to {{text|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set newShape to make new shape with properties {auto shape type:{{type|autoshape}}, left position:leftPos, top:topPos, width:widthVal, height:heightVal}
            if txtArg is not "" then
                try
                    set content of text range of text frame of newShape to txtArg
                end try
            end if
            return (count of shapes) as string
        end tell
    end tell
end tell
