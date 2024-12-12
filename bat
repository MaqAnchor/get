@echo off
set USERNAME=your_username
set PASSWORD=your_password

for /f "usebackq tokens=1" %%i in ("computers.txt") do (
    echo Running commands on %%i...
    psexec \\%%i -u %USERNAME% -p %PASSWORD% cmd /c "netstat -a -n -o | find \"80\" && netstat -a -n -o | find \"8080\" && tasklist" > %%i_output.txt 2>&1
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to connect to %%i. Skipping...
        rem Optionally delete the output file if it's empty or contains just an error message
        del %%i_output.txt >nul 2>&1
    ) else (
        echo Output for %%i saved to %%i_output.txt
    )
)

echo All processing complete.
pause
