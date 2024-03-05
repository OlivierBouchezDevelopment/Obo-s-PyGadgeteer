
REM In visual sudio code , if you can execute this directly with the "RUN ARROW" install batch runner extension
REM https://marketplace.visualstudio.com/items?itemName=NilsSoderman.batch-runner

REM CREATE THE DOCS
REM It is written for Visual Studio Code
REM %WORKSPACEFOLDER% is the root of the workspace

CALL workon Obo-s-PyGadgeteer
IF %ERRORLEVEL% NEQ 0 (
    ECHO Failed to activate virtual environment.
    EXIT /B %ERRORLEVEL%
)
cd %WORKSPACEFOLDER%
REM to create conf.py and index.rst and the structure
REM sphinx-quickstart docs

cd %WORKSPACEFOLDER%
sphinx-apidoc  --force -o docs/source pygadgeteer

cd %WORKSPACEFOLDER%
sphinx-build -b html docs/source dist/docs