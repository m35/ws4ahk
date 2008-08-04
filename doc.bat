SET VER=0.21

robodoc-4-99-36.exe --rc ws.rc --documenttitle "Windows Scripting for Autohotkey v%VER% Public API" --doc "Windows Scripting for Autohotkey v%VER% Public API" 

copy "Windows Scripting for Autohotkey v%VER% Public API.html" ws4ahk_public_api.html

robodoc-4-99-36.exe --rc ws.rc --documenttitle "Windows Scripting for Autohotkey v%VER% Internal API" --internalonly --doc "Windows Scripting for Autohotkey v%VER% Internal API"

copy "Windows Scripting for Autohotkey v%VER% Internal API.html" ws4ahk_internal_api.html

copy "Windows Scripting for Autohotkey v%VER% Public API.css" ws4ahk.css

@echo Now edit these html files with the following:
@echo * Change styleguide link
@echo * Get rid of path in the title
