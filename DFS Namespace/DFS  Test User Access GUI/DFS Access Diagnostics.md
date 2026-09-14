## Original PowerShell Code by Radu Vuia!!! Thanks Radu!
# GUI Wrapper created by Stephen McKee
# DFS Access Diagnostics - GUI wrapper
# Wraps the DFS/SMB diagnostic logic in a WinForms front end.
# Run with:  powershell.exe -STA -File .\DFS-Diagnostics-GUI.ps1

## DFS Diagnostics GUI Walkthrough
    The window gives you a clean way to launch the diagnostic, monitor progress, and review outcomes without touching the command line.
    UNC path input – Pre-filled with your original DFS path, but you can change it to any \server\share\folder target.
    Report output folder – Choose where the timestamped report and CSV summary are saved; use Browse to pick a location.
    Write access toggle – Tick the checkbox to test creating and deleting a small temp file in the share, which is otherwise skipped.
    Run button – Disables the controls, starts the background run-space, and streams each check result into the output box with color coding (green PASS, red FAIL, orange WARN, blue STEP).
    Action buttons – Open the last report folder, copy the output text, or clear the output area. A status bar at the bottom shows elapsed time and final failure/warning counts.
    Optimization Tip: You can adjust the default UNC path in the $tbPath.Text line and the output folder in the $tbOut.Text line near the top of the GUI section to match your environment. 
    The background polling interval is set by $timer.Interval = 150 (milliseconds).

