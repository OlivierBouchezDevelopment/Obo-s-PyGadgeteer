
REM START A WEBSERVER TO VIEW THE DOCS
REM It is written for Visual Studio Code
REM %WORKSPACEFOLDER% is the root of the workspace

start "" http://localhost:8001
python -m http.server -d "%WORKSPACEFOLDER%/dist/docs/" 8001