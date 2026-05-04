-- ppt notes.set <slide_index> --text=<notes_text>
set slideIdx to ({{slide|json}} as integer)
set notesText to {{text|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            -- Each slide exposes its speaker notes as `notes page`, whose
            -- shape 2 is the notes placeholder (shape 1 is the slide image).
            set notesShape to shape 2 of notes page
            set content of text range of text frame of notesShape to notesText
            return "updated"
        end tell
    end tell
end tell
