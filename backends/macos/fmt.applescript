-- ppt fmt <slide_index> --shape=<idx> [--bold] [--italic] [--size]
--                                      [--color] [--bg] [--align]
--
-- Each attribute is wrapped in (a) {{#if name}}…{{/if}} so unspecified
-- attributes drop out at template-render time, and (b) a try/on-error
-- so a bad value for one doesn't abort the others.
-- Returns "okList|failList" comma-separated, mirroring the Swift bridge.
set slideIdx to ({{slide|json}} as integer)
set shapeIdx to ({{shape|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set sh to shape shapeIdx
            set tr to text range of text frame of sh
            set okList to {}
            set failList to {}

            {{#if bold}}
            try
                set bold of font of tr to true
                set end of okList to "bold"
            on error errMsg
                set end of failList to "bold: " & errMsg
            end try
            {{/if}}

            {{#if italic}}
            try
                set italic of font of tr to true
                set end of okList to "italic"
            on error errMsg
                set end of failList to "italic: " & errMsg
            end try
            {{/if}}

            {{#if size}}
            try
                set font size of font of tr to ({{size|json}} as real)
                set end of okList to "size"
            on error errMsg
                set end of failList to "size: " & errMsg
            end try
            {{/if}}

            {{#if color}}
            try
                set font color of font of tr to {{color|rgb}}
                set end of okList to "color"
            on error errMsg
                set end of failList to "color: " & errMsg
            end try
            {{/if}}

            {{#if bg}}
            try
                set fore color of fill format of sh to {{bg|rgb}}
                set end of okList to "bg"
            on error errMsg
                set end of failList to "bg: " & errMsg
            end try
            {{/if}}

            {{#if align}}
            try
                set paragraph alignment of paragraph format of tr to {{align|align}}
                set end of okList to "align"
            on error errMsg
                set end of failList to "align: " & errMsg
            end try
            {{/if}}

            set AppleScript's text item delimiters to ","
            set okStr to okList as string
            set failStr to failList as string
            set AppleScript's text item delimiters to ""
            return okStr & "|" & failStr
        end tell
    end tell
end tell
