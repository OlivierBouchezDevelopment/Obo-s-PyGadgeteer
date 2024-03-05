
CALL workon Obo-s-PyGadgeteer
IF %ERRORLEVEL% NEQ 0 (
    ECHO Failed to activate virtual environment.
    EXIT /B %ERRORLEVEL%
)
python -m build