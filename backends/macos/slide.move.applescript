-- ppt slide.move <from> --to=<to>
set fromIdx to ({{from|json}} as integer)
set toIdx to ({{to|json}} as integer)

tell application "Microsoft PowerPoint"
    tell active presentation
        move slide fromIdx to before slide toIdx
        return (count of slides) as string
    end tell
end tell
