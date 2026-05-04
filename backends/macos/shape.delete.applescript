-- ppt shape.delete <slide_index> --shape=<idx>
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            delete shape shapeIdx
            return (count of shapes) as string
        end tell
    end tell
end tell
