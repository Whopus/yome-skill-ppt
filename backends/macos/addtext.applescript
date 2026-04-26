-- ppt addtext <slide_index> --text=<text> [--left] [--top] [--width] [--height]
--                                          [--size] [--color] [--bold] [--italic]
--
-- Each optional formatting attribute is wrapped in {{#if name}}…{{/if}} so it
-- only renders when actually given, then in a per-attribute try/on-error so
-- one bad value can't kill the whole creation.
set slideIdx to ({{slide|json}} as integer)
set newText to {{text|json}}
set leftPos to ({{left|json}} as real)
set topPos to ({{top|json}} as real)
set widthVal to ({{width|json}} as real)
set heightVal to ({{height|json}} as real)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set newShape to make new shape at end with properties {left position:leftPos, top:topPos, width:widthVal, height:heightVal}
            set content of text range of text frame of newShape to newText

            {{#if size}}
            try
                set font size of font of text range of text frame of newShape to ({{size|json}} as real)
            end try
            {{/if}}

            {{#if color}}
            try
                set font color of font of text range of text frame of newShape to {{color|rgb}}
            end try
            {{/if}}

            {{#if bold}}
            try
                set bold of font of text range of text frame of newShape to true
            end try
            {{/if}}

            {{#if italic}}
            try
                set italic of font of text range of text frame of newShape to true
            end try
            {{/if}}

            return (count of shapes) as string
        end tell
    end tell
end tell
