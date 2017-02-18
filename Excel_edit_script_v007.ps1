try {
    # SDG. Anton Thelander.

    # Released under GNU/GPL v3.0 or later.
    # No express or implied warranties for this program and subprograms.
    
    $Now = Get-Date -UFormat "%Y-%m-%d %H:%M:%S";
    function GetNow {
        $TempNow = Get-Date -UFormat "%Y-%m-%d %H:%M:%S";
        $Now = $TempNow;
    }

    GetNow; Write-Host "$Now. Please target Genesis or another book in Excel before running this script.";
    GetNow; Write-Host "$Now. WARNING! IF YOU DON'T KNOW WHAT THIS SCRIPT IS DOING";
    GetNow; Write-Host "$Now. AND ARE NOT AWARE OF THE RISKS PLEASE CLOSE THIS WINDOW,";
    GetNow; Write-Host "$Now. OR PRESS CONTROL+C INSIDE OF THE POWERSHELL SCRIPT.";
    Pause;
    
    while (1) {
        Sleep -m 30;
        $Roof = 0;

        # SendKeys: + = shift, ^ = ctrl, % = win.

        # Escape.
        Select-Window *excel* | Send-Keys "{ESC}";
        Sleep -m 30;

        # Refresh.
        Select-Window *excel* | Send-Keys "{RIGHT}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "{LEFT}";
        Sleep -m 30;

        # Extract book name.
        Select-Window *excel* | Send-Keys "{F2}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "^+{HOME}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "^{c}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "{ESC}";

        $Book = Get-Clipboard; $Book = [string]$Book;

        # Acquire the roof/chapter number.
        Select-Window *excel* | Send-Keys "{RIGHT}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "{F2}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "^+{HOME}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "^{c}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "{ESC}";
        Sleep -m 30;
        Select-Window *excel* | Send-Keys "{LEFT}";
        Sleep -m 30;

        $Roof = Get-Clipboard; $Roof = [int]$Roof;
    
        if ( -not ( ( $Roof -gt 0 ) -and ( $Roof -lt 151 ) ) ) {Exit-PSHostProcess;};

        GetNow; Write-Host "$Now. Book: $Book.";
        GetNow; Write-Host "$Now. Roof/Chapters: $Roof.";

        if ( $Roof -gt 1 ) {

            # Add rows.
            Select-Window *excel* | Send-Keys "^+{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^+{RIGHT}";
            Sleep -m 30;
            if ( $Roof -gt 2 ) {
                $X = 3; $X = [int]$X;
                while ($X -le $Roof) {Select-Window *excel* | Send-Keys "+{DOWN}"; Sleep -m 35; $X = $X + 1;};
            }

            Select-Window *excel* | Send-Keys "^{+}";
            Sleep -m 30;
    
            # Copy and paste. Chapters.
            Select-Window *excel* | Send-Keys "{ESC}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "1";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{TAB}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "+{TAB}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{c}";
            Sleep -m 30;
            if ( $Roof -gt 2 ) {
                Select-Window *excel* | Send-Keys "+^{DOWN}";
                Sleep -m 30;
            } else {
                Select-Window *excel* | Send-Keys "+{DOWN}";
                Sleep -m 30;
            }
            Select-Window *excel* | Send-Keys "^{v}";
            Sleep -m 100;
            Select-Window *excel* | Send-Keys "%{h}";
            Sleep -m 100;
            Select-Window *excel* | Send-Keys "{f}";
            Sleep -m 100;
            Select-Window *excel* | Send-Keys "{i}";
            Sleep -m 200;
            Select-Window *excel* | Select-ChildWindow | Send-Keys "{s}";
            Sleep -m 700;
            Select-Window *excel* | Select-ChildWindow | Send-Keys "{ENTER}";
            Sleep -m 100;

            # Move and make a mark for the next book.
            Select-Window *excel* | Send-Keys "{LEFT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{DOWN}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{x}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{LEFT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{LEFT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{UP}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{DOWN}";
            Sleep -m 30;

            # Book name.
            Set-Clipboard $Book;
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{v}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{c}";
            Sleep -m 30;
            if ( $Roof -gt 2 ) {
                Select-Window *excel* | Send-Keys "^+{DOWN}";
                Sleep -m 30;
            } else {
                Select-Window *excel* | Send-Keys "+{DOWN}";
                Sleep -m 30;
            }
            Select-Window *excel* | Send-Keys "^{v}";
            Sleep -m 30;

            # Move to the next book.
            Select-Window *excel* | Send-Keys "{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{RIGHT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "^{DOWN}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{DEL}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{DOWN}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{LEFT}";
            Sleep -m 30;
            Select-Window *excel* | Send-Keys "{LEFT}";
            Sleep -m 30;

            #Pause;

            Sleep -m 300;
        } else {
            Select-Window *excel* | Send-Keys "{DOWN}";
            Sleep -m 30;
        }<# else {
            GetNow; Write-Host "$Now. Bad number in roof/chapters. Exiting.";
            Pause;
            Exit-PSHostProcess;
        }#>
    }

    Select-Window *ise* | Send-Keys;

}
catch {

}