-- ppt notes.get <slide_index>
set slideIdx to ({{slide|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        tell slide slideIdx
            set notesText to ""
            try
                set notesShape to shape 2 of notes page
                set notesText to content of text range of text frame of notesShape
            end try
            return notesText
        end tell
    end tell
end tell
