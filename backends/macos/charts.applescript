-- ppt charts <slide_index>
-- Returns TSV: shape<tab>type<tab>title<tab>seriesCount for every chart shape on the slide.
set slideIdx to ({{slide|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set shapeCount to count of shapes
            set lineList to {}
            repeat with i from 1 to shapeCount
                set sh to shape i
                set isChart to false
                try
                    set isChart to has chart of sh
                end try
                if isChart then
                    set chartType to ""
                    set chartTitle to ""
                    set seriesCount to 0
                    try
                        set chartType to (chart type of chart of sh) as string
                    end try
                    try
                        set chartTitle to caption of chart title of chart of sh
                    end try
                    try
                        set seriesCount to count of series of chart of sh
                    end try
                    set end of lineList to (i as string) & tab & chartType & tab & chartTitle & tab & (seriesCount as string)
                end if
            end repeat
            set AppleScript's text item delimiters to linefeed
            return lineList as string
        end tell
    end tell
end tell
