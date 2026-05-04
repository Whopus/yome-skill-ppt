-- ppt chart.delete <slide_index> --shape=<n>
-- Refuses to delete a non-chart shape to avoid accidental data loss.
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set sh to shape shapeIdx
            set isChart to false
            try
                set isChart to has chart of sh
            end try
            if not isChart then
                error "shape " & shapeIdx & " on slide " & slideIdx & " is not a chart" number 7102
            end if
            delete sh
            return (count of shapes) as string
        end tell
    end tell
end tell
