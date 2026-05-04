-- ppt table.set <slide_index> --shape=<idx> --row=r --col=c --text=<cell>
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)
set rowIdx to ({{row|json}} as integer)
set colIdx to ({{col|json}} as integer)
set cellText to {{text|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set tblShape to shape shapeIdx
            set tblObj to table object of tblShape
            set theCell to get cell from tblObj row rowIdx column colIdx
            -- `get cell from` returns a cell whose shape property is the
            -- underlying textbox-like shape of that cell.
            set cellShape to shape of theCell
            set content of text range of text frame of cellShape to cellText
            return "updated"
        end tell
    end tell
end tell
