-- ppt chart.add <slide_index> --data=<CSV> [--type] [--title] [--left] [--top] [--width] [--height]
-- CSV format:
--   Row 1: header — first cell is category label (may be empty), remaining cells are series names.
--   Row 2+: "<category>,<v1>[,<v2>...]"
-- Newlines may arrive as literal "\n" inside the --data argument.
set slideIdx to ({{slide|json}} as integer)
set rawData to {{data|json}}
set chartKind to {{type|json}}
set chartTitle to {{title|json}}
set leftPos to ({{left|json}} as real)
set topPos to ({{top|json}} as real)
set widthVal to ({{width|json}} as real)
set heightVal to ({{height|json}} as real)

-- Normalise "\n" escape sequences produced by the CLI to real line feeds.
set AppleScript's text item delimiters to "\\n"
set parts to text items of rawData
set AppleScript's text item delimiters to linefeed
set rawData to parts as string
set AppleScript's text item delimiters to ""

-- Split rows
set rowList to paragraphs of rawData
if (count of rowList) < 2 then
    error "ppt chart.add: --data needs at least 2 rows (header + 1 data row)" number 7100
end if

-- Map type → AppleScript chart type constant
set ptType to column clustered
if chartKind is "bar" then
    set ptType to bar clustered
else if chartKind is "line" then
    set ptType to line
else if chartKind is "pie" then
    set ptType to pie
else if chartKind is "area" then
    set ptType to area
else if chartKind is "scatter" then
    set ptType to XY scatter
end if

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set newShape to add chart chart type ptType with properties {left position:leftPos, top:topPos, width:widthVal, height:heightVal}
            set theChart to chart of newShape

            -- Title (optional, wrapped in try so an unsupported object model does not break creation)
            if chartTitle is not "" then
                try
                    set has title of theChart to true
                    set caption of chart title of theChart to chartTitle
                end try
            end if

            -- Feed data via the chart data sheet. Values default to 0 on parse failure
            -- so a single bad cell does not abort the whole chart.
            try
                tell data sheet of chart data of theChart
                    -- Clear existing default sample data
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
                            -- Header row: keep strings. Data rows column 1: category label. Others: numeric.
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
