-- ppt slide.duplicate <slide_index> [--to=<target>]
-- Without --to the copy is inserted right after the source.
set slideIdx to ({{slide|json}} as integer)
set toArg to {{to|json}}

tell application "Microsoft PowerPoint"
    tell active presentation
        set origSlide to slide slideIdx
        set dupSlide to duplicate origSlide
        set sCount to count of slides
        -- `duplicate` returns the new slide which is placed at the end by default.
        -- Compute final destination: provided --to index, or (source+1).
        set dstIdx to sCount
        if toArg is not "" then
            set dstIdx to (toArg as integer)
        else
            set dstIdx to slideIdx + 1
        end if
        if dstIdx < 1 then set dstIdx to 1
        if dstIdx > sCount then set dstIdx to sCount
        -- Determine current index of dup (last slide) and move if needed.
        if dstIdx is not sCount then
            move slide sCount to before slide dstIdx
        end if
        return "duplicated"
    end tell
end tell
