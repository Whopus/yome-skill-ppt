-- ppt read <slide_index> --shape=<shape_index>
-- Returns TSV (single line):
-- name<tab>text<tab>bold<tab>italic<tab>fontSize<tab>fontName<tab>colorRGB
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set sh to shape shapeIdx
            set shName to name of sh
            set hasText to has text frame of sh
            set textContent to ""
            set isBold to ""
            set isItalic to ""
            set fSize to ""
            set fName to ""
            set fColorStr to ""
            if hasText then
                try
                    set textContent to content of text range of text frame of sh
                    set tr to text range of text frame of sh
                    set isBold to bold of font of tr as string
                    set isItalic to italic of font of tr as string
                    set fSize to font size of font of tr as string
                    set fName to font name of font of tr
                    set fColor to font color of font of tr
                    set fColorStr to (item 1 of fColor as string) & "," & (item 2 of fColor as string) & "," & (item 3 of fColor as string)
                end try
            end if
            return shName & tab & textContent & tab & isBold & tab & isItalic & tab & fSize & tab & fName & tab & fColorStr
        end tell
    end tell
end tell
