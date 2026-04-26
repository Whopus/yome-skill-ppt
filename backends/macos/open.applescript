-- ppt open <path>
-- Triggered via openViaLaunchServices in dispatcher; this template is
-- only the fallback. The poll script (open.poll.applescript) is what
-- the dispatcher actually polls after Launch Services hands the file
-- to PowerPoint.
tell application "Microsoft PowerPoint"
    activate
    open {{path|posix}}
    delay 0.5
    return name of active presentation
end tell
