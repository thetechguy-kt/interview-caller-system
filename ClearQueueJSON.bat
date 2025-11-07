@echo off
set FILE=queue_state.json

:: Delete the file if it exists
if exist "%FILE%" del "%FILE%"

:: Create a new file with default JSON
(
echo {
echo     "queue": [],
echo     "called_tokens": []
echo }
) > "%FILE%"

echo %FILE% has been reset successfully.
pause
