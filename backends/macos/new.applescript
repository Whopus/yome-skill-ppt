-- ppt new [<path>] [--force]
-- Creates a new blank presentation. If path is given, also save as that path
-- using Open XML (.pptx) format.
tell application "Microsoft PowerPoint"
    activate
    delay 1
    make new presentation
    delay 0.5

    -- `make new presentation` produces a zero-slide deck on macOS PowerPoint;
    -- seed one slide so `ppt new` leaves the deck in an editable state.
    -- The `tell active presentation` context is what makes `make new slide
    -- at end` work (outside that context PowerPoint can't generate the
    -- class — error -2710). We do NOT touch `layout` here: `slide layout N`
    -- is a syntax error on some PowerPoint versions; downstream `slide.add`
    -- can set layouts when the user explicitly asks.
    try
        tell active presentation
            make new slide at end
        end tell
    end try

    delay 0.3

    {{#if path}}
    set destPath to POSIX file {{path|json}} as string
    tell active presentation
        save in destPath as save as Open XML presentation
    end tell
    {{/if}}
    return name of active presentation
end tell
