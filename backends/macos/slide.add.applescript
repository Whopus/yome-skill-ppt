-- ppt slide.add [--index=<n>] [--layout=<n|name>]
-- Without --index: append at end. With --index=N: insert before slide N.
-- --layout: numeric index OR layout name; numeric is more reliable.
set posArg to {{index|json}}
set layoutArg to {{layout|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        if posArg is "" then
            set createdSlide to make new slide at end
        else
            set targetIdx to (posArg as integer)
            set createdSlide to make new slide at before slide targetIdx
        end if

        set sCount to count of slides
    end tell
end tell

-- Layout setting is intentionally OUTSIDE the tell-block above and is
-- evaluated through `run script` so the (version-dependent)
-- `slide layout N` term is only compiled when the user actually asked
-- for a layout. On PowerPoint builds where that term doesn't exist the
-- whole template would otherwise refuse to compile (error -10002).
if layoutArg is not "" then
    try
        set scriptSrc to "tell application \"Microsoft PowerPoint\" to set layout of slide " & sCount & " of active presentation to slide layout " & layoutArg & " of slide master 1 of active presentation"
        run script scriptSrc
    end try
end if

return sCount as string
