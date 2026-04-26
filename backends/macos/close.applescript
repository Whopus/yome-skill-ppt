-- ppt close [--save=true|false]
set shouldSave to {{save|bool}}
tell application "Microsoft PowerPoint"
    if shouldSave then
        close active presentation saving yes
    else
        close active presentation saving no
    end if
end tell
return "closed"
