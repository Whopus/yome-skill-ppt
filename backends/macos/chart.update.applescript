-- ppt chart.update <slide_index> --shape=<n> --data=<CSV>
-- Replace strategy: capture the existing chart's type / position / title, delete it,
-- then recreate at the same spot with the new CSV. Safer than in-place data sheet edits
-- across PowerPoint versions.
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)
set rawData to {{data|json}}

set AppleScript's text item delimiters to "\\n"
set parts to text items of rawData
set AppleScript's text item delimiters to linefeed
set rawData to parts as string
set AppleScript's text item delimiters to ""

set rowList to paragraphs of rawData
if (count of rowList) < 2 then
    error "ppt chart.update: --data needs at least 2 rows (header + 1 data row)" number 7101
end if

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            -- Snapshot existing chart metadata
            set origShape to shape shapeIdx
            set origChart to chart of origShape
            set origType to chart type of origChart
            set origLeft to left position of origShape
            set origTop to top of origShape
            set origWidth to width of origShape
            set origHeight to height of origShape
            set origTitle to ""
            try
                set origTitle to caption of chart title of origChart
            end try

            delete origShape

            set newShape to add chart chart type origType with properties {left position:origLeft, top:origTop, width:origWidth, height:origHeight}
            set theChart to chart of newShape

            if origTitle is not "" then
                try
                    set has title of theChart to true
                    set caption of chart title of theChart to origTitle
                end try
            end if

            try
                tell data sheet of chart data of theChart
                    try
                        clear contents of cell range "A1:Z100"
                    end try

                    set rowCount to count of rowList
                    repeat with r from 1 to rowCount
                        set rowText to item r of rowList
                        set AppleScript's text item delimiters to ","
                        set cells to text items of rowText
                        set AppleScript's text item delimiters to ""
                        set colCount to count of cells
                        repeat with c from 1 to colCount
                            set cellVal to item c of cells
                            if r is 1 or c is 1 then
                                set value of cell (r as string) & (character id (64 + c)) to cellVal
                            else
                                try
                                    set value of cell (r as string) & (character id (64 + c)) to (cellVal as real)
                                on error
                                    set value of cell (r as string) & (character id (64 + c)) to 0
                                end try
                            end if
                        end repeat
                    end repeat
                end tell
            end try

            return (count of shapes) as string
        end tell
    end tell
end tell
