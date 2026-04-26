-- ppt slide.delete <index> [--force]
-- Refuses to delete a non-empty slide unless --force is given.
set idx to ({{index|json}} as integer)
set forceFlag to {{force|bool}}

tell application "Microsoft PowerPoint"
    tell active presentation
        if not forceFlag then
            set shapeCount to count of shapes of slide idx
            if shapeCount > 0 then
                error "slide " & idx & " has " & shapeCount & " shape(s); pass --force to delete non-empty slide" number 7000
            end if
        end if
        delete slide idx
        return (count of slides) as string
    end tell
end tell
