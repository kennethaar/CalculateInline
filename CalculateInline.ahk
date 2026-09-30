#Requires AutoHotkey v2.0
#SingleInstance Force
/*
FEATURES:
Nothing selected:
- Hotkey opens a simple calculator dialog

If mathematical expression is selected:
- Replaces selected mathematical expression with calculation results
- To undo press Ctrl+Z

If numbers are selected:
- A results window shows sum, count, average (mean), and median
- Copy any single value, or all of them, with one button - nothing is
  copied until you press one of those buttons
- Shows the Net Promoter Score (NPS) when every value is a 0-10 score
- Press S, A, M, C, N or 4 to copy sum, average, median, count, NPS or all
  values and close the window in one keystroke
- Support for numbers with thousand separators
- All values are displayed with two decimal places

Settings:
- Configurable hotkey via tray menu
- Toggle between European/American number format

Licenses:
- Script released under the GPL v3
- Icon from pictogrammes.com
- Ahk library released under the GPL v2

To trigger the app press
*/

; Define script variables
global VarScriptName := "CalculateInline"
global VarVersionNo := "v21"
global ConfigFile := A_ScriptDir "\CalculateInline.ini"
global CurrentHotkey := "!c"  ; Default hotkey (Alt+C)
global UseEuropeanFormat := true  ; Default to European format (comma as decimal separator)
global originalSelectedText := ""  ; Store original selected text for recalculation
global ActiveResultsGui := ""     ; The results window while one is open
global ActiveResultsAction := ""  ; Its "copy this and close" handler, driven by the shortcut keys

; ===== Function Definitions =====

; Function to calculate median of an array of numbers
CalculateMedian(numbers) {
    if numbers.Length = 0
        return 0

    ; Create a copy of the array to sort
    sortedNumbers := []
    for num in numbers
        sortedNumbers.Push(num)

    ; Sort the numbers
    n := sortedNumbers.Length
    loop n {
        i := A_Index
        loop n - i {
            j := A_Index + i
            if (sortedNumbers[j] < sortedNumbers[j-1]) {
                temp := sortedNumbers[j]
                sortedNumbers[j] := sortedNumbers[j-1]
                sortedNumbers[j-1] := temp
            }
        }
    }

    ; Calculate median
    if Mod(n, 2) = 1 {
        ; Odd number of elements - median is the middle one
        return sortedNumbers[Ceil(n/2)]
    } else {
        ; Even number of elements - median is average of two middle elements
        return (sortedNumbers[n/2] + sortedNumbers[n/2 + 1]) / 2
    }
}

; Simplified message box function
ShowMessage(text, title := "", options := "OK") {
    title := title ? title : VarScriptName " " VarVersionNo
    return MsgBox(text, title, options)
}

; Function to show calculator input box
ShowCalculator(*) {
    expressionInput := InputBox("Enter a mathematical expression to be calculated:",
                               VarScriptName " " VarVersionNo, "w300 h150")

    if expressionInput.Result = "Cancel"
        return

    expression := expressionInput.Value

    try {
        result := CalculateExpression(expression)
        A_Clipboard := result
        ShowMessage("Result: " result "`n`nCopied to clipboard")
    } catch as err {
        ShowMessage("Error calculating: " expression "`n" err.Message)
    }
}

; Main function to handle calculation operations
CalculateSelectedTextOrShowCalculator(*) {
    savedClipboard := A_Clipboard
    A_Clipboard := ""  ; Clear clipboard

    Send "^c"
    if !ClipWait(0.75) {
        A_Clipboard := savedClipboard
        ShowCalculator()
        return
    }

    selectedText := A_Clipboard
    global originalSelectedText := selectedText  ; Store in a global variable

    if IsCalculatableExpression(selectedText) {
        ProcessExpression(selectedText, savedClipboard)
    } else if ContainsNumbers(selectedText) {
        ProcessMultipleNumbers(selectedText, savedClipboard)
    } else {
        A_Clipboard := savedClipboard
        ShowCalculator()
    }
}

; Function to check if text contains any numbers
ContainsNumbers(text) {
    return RegExMatch(text, "\d")
}

; Function to process text as a mathematical expression
ProcessExpression(text, savedClipboard) {
    try {
        result := CalculateExpression(text)
        A_Clipboard := result
        Send "^v"

        ShowMessage("Calculated: " text " = " result "`n`nResult has been copied to clipboard and replaced the selected text.")
    } catch as err {
        A_Clipboard := savedClipboard
        ShowMessage("Error calculating: " text "`n" err.Message)
        ShowCalculator()
    }
}

; Build the set of statistics for an array of numbers
BuildStats(numbers) {
    sum := 0
    for number in numbers
        sum += number

    return {count: numbers.Length,
            sum: sum,
            average: numbers.Length > 0 ? sum / numbers.Length : 0,
            median: CalculateMedian(numbers),
            nps: CalculateNps(numbers)}
}

; Net Promoter Score: % promoters (9-10) minus % detractors (0-6), from -100 to 100.
; Only meaningful when every value is a whole-number score from 0 to 10, so any
; other selection returns "" and the results window shows NPS as not available.
CalculateNps(numbers) {
    if numbers.Length = 0
        return ""

    promoters := 0, passives := 0, detractors := 0
    for number in numbers {
        if (number < 0 || number > 10 || number != Round(number))
            return ""
        if (number >= 9)
            promoters++
        else if (number >= 7)
            passives++
        else
            detractors++
    }

    return {score: Round((promoters - detractors) / numbers.Length * 100),
            promoters: promoters,
            passives: passives,
            detractors: detractors}
}

; Function to process text as multiple numbers for sum, count, average and median
ProcessMultipleNumbers(text, savedClipboard) {
    numbers := ExtractNumbers(text)

    if numbers.Length = 0 {
        A_Clipboard := savedClipboard
        ShowMessage("No numbers found in selection")
        return
    }

    ; Nothing is copied unless the user asks for it, so restore the clipboard now
    A_Clipboard := savedClipboard

    ShowResultsWindow(BuildStats(numbers), BuildStats(ExtractNumbersWithSeparators(text)))
}

; Results window - each button says exactly what it copies, and the clipboard
; is only touched when one of them is pressed
ShowResultsWindow(plainStats, sepStats) {
    global ActiveResultsGui, ActiveResultsAction

    ; Only one results window at a time, so the shortcut keys are never ambiguous
    if (ActiveResultsGui) {
        try ActiveResultsGui.Destroy()
        ActiveResultsGui := ""
        ActiveResultsAction := ""
    }

    ; Only offer the two readings when they actually disagree
    hasSeparators := sepStats.count > 0
                     && (sepStats.count != plainStats.count || sepStats.sum != plainStats.sum)

    resultsGui := Gui("+AlwaysOnTop -MinimizeBox", VarScriptName " " VarVersionNo " - Results")
    resultsGui.SetFont("s10")

    if (hasSeparators) {
        separatorExample := UseEuropeanFormat ? "1.234,56" : "1,234.56"
        resultsGui.Add("Text", "xm w355", "The selection can be read two ways:")
        plainRadio := resultsGui.Add("Radio", "xm y+4 w355 Checked",
                                     "Plain numbers - " plainStats.count " values found")
        sepRadio := resultsGui.Add("Radio", "xm y+4 w355",
                                   "With thousand separators (" separatorExample ") - " sepStats.count " values found")
        resultsGui.Add("Text", "xm y+10 w355 0x10")  ; Horizontal divider
    }

    resultsGui.Add("Text", "xm y+10 w110", "Count:")
    countValue := resultsGui.Add("Text", "x+5 yp w240", "")
    resultsGui.Add("Text", "xm y+6 w110", "Sum (total):")
    sumValue := resultsGui.Add("Text", "x+5 yp w240", "")
    resultsGui.Add("Text", "xm y+6 w110", "Average (mean):")
    averageValue := resultsGui.Add("Text", "x+5 yp w240", "")
    resultsGui.Add("Text", "xm y+6 w110", "Median:")
    medianValue := resultsGui.Add("Text", "x+5 yp w240", "")
    resultsGui.Add("Text", "xm y+6 w110", "NPS:")
    npsValue := resultsGui.Add("Text", "x+5 yp w240", "")
    npsDetail := resultsGui.Add("Text", "xp y+2 w240", "")
    npsDetail.SetFont("s9")

    for valueControl in [countValue, sumValue, averageValue, medianValue, npsValue]
        valueControl.SetFont("s10 Bold")

    resultsGui.Add("Text", "xm y+14 w355", "Copy to clipboard:")
    sumButton := resultsGui.Add("Button", "xm y+4 w85", "Sum (S)")
    averageButton := resultsGui.Add("Button", "x+5 yp w85", "Average (A)")
    medianButton := resultsGui.Add("Button", "x+5 yp w85", "Median (M)")
    countButton := resultsGui.Add("Button", "x+5 yp w85", "Count (C)")
    allButton := resultsGui.Add("Button", "xm y+5 w175", "All values on one line (4)")
    npsButton := resultsGui.Add("Button", "x+5 yp w85", "NPS (N)")
    closeButton := resultsGui.Add("Button", "x+5 yp w85 Default", "Close")

    resultsGui.Add("Text", "xm y+10 w355",
                   "Pressing S, A, M, C, N or 4 copies that value and closes this window. Clicking a button copies and leaves it open.")
    statusText := resultsGui.Add("Text", "xm y+6 w355", "Clipboard not changed yet.")

    if (hasSeparators) {
        plainRadio.OnEvent("Click", UpdateDisplay)
        sepRadio.OnEvent("Click", UpdateDisplay)
    }
    sumButton.OnEvent("Click", (*) => CopyValue("sum", FormatNumber(CurrentStats().sum)))
    averageButton.OnEvent("Click", (*) => CopyValue("average", FormatNumber(CurrentStats().average)))
    medianButton.OnEvent("Click", (*) => CopyValue("median", FormatNumber(CurrentStats().median)))
    countButton.OnEvent("Click", (*) => CopyValue("count", CurrentStats().count))
    npsButton.OnEvent("Click", (*) => CopyValue("NPS", CurrentStats().nps.score))
    allButton.OnEvent("Click", CopyAllValues)
    closeButton.OnEvent("Click", CloseWindow)
    resultsGui.OnEvent("Escape", CloseWindow)
    resultsGui.OnEvent("Close", CloseWindow)

    UpdateDisplay()
    resultsGui.Show("AutoSize Center")

    ; Arm the single-key shortcuts for this window
    ActiveResultsGui := resultsGui
    ActiveResultsAction := CopyAndClose

    ; Copy one value and close - what the S/A/M/C/N/4 keys do
    CopyAndClose(which) {
        ; No NPS for this selection - say so and leave the window open
        if (which = "nps" && !CurrentStats().nps) {
            ToolTip("NPS needs every value to be a whole-number score from 0 to 10")
            SetTimer () => ToolTip(), -2000
            return
        }

        switch which {
            case "sum":     CopyValue("sum", FormatNumber(CurrentStats().sum))
            case "average": CopyValue("average", FormatNumber(CurrentStats().average))
            case "median":  CopyValue("median", FormatNumber(CurrentStats().median))
            case "count":   CopyValue("count", CurrentStats().count)
            case "nps":     CopyValue("NPS", CurrentStats().nps.score)
            case "all":     CopyAllValues()
        }

        ; The window is about to vanish, so confirm the copy with a brief tooltip
        ToolTip(statusText.Value)
        SetTimer () => ToolTip(), -2000

        CloseWindow()
    }

    CloseWindow(*) {
        global ActiveResultsGui, ActiveResultsAction
        ActiveResultsGui := ""
        ActiveResultsAction := ""
        resultsGui.Destroy()
    }

    ; The statistics the window is currently showing
    CurrentStats() {
        return (hasSeparators && sepRadio.Value) ? sepStats : plainStats
    }

    UpdateDisplay(*) {
        stats := CurrentStats()
        countValue.Value := stats.count
        sumValue.Value := FormatNumber(stats.sum)
        averageValue.Value := FormatNumber(stats.average)
        medianValue.Value := FormatNumber(stats.median)

        if (stats.nps) {
            npsValue.Value := stats.nps.score
            npsDetail.Value := stats.nps.promoters " promoters, " stats.nps.passives " passives, "
                             . stats.nps.detractors " detractors"
        } else {
            npsValue.Value := "n/a"
            npsDetail.Value := "Needs whole-number scores from 0 to 10"
        }
        npsButton.Enabled := !!stats.nps
    }

    CopyValue(label, value) {
        A_Clipboard := value
        statusText.Value := "Copied " label " to clipboard: " value
    }

    CopyAllValues(*) {
        stats := CurrentStats()
        A_Clipboard := "Sum: " FormatNumber(stats.sum) ", Count: " stats.count
                    . ", Average: " FormatNumber(stats.average) ", Median: " FormatNumber(stats.median)
        if (stats.nps)
            A_Clipboard .= ", NPS: " stats.nps.score
        statusText.Value := "Copied all values to clipboard."
    }
}

; Single-key shortcuts that are only live while the results window is focused.
; Registered once at startup and matched on the window title, so repeated use of
; the script never piles up hotkey variants.
RegisterResultsShortcuts() {
    HotIfWinActive(VarScriptName " " VarVersionNo " - Results")
    for key, action in Map("s", "sum", "a", "average", "m", "median", "c", "count", "n", "nps", "4", "all")
        Hotkey key, ResultsShortcut.Bind(action), "On"
    HotIf()
}

ResultsShortcut(action, *) {
    if (ActiveResultsAction)
        ActiveResultsAction.Call(action)
}

; Combined function to extract numbers with thousand separators.
; Every line is parsed exactly once, so no value can be counted twice.
ExtractNumbersWithSeparators(text) {
    numbers := []

    ; Matches the longest form first, so "1.234.567,89" is one number and not two
    scanPattern := UseEuropeanFormat ? "\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+" : "\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+"

    ; Only the active format's decimal sign counts as a decimal sign here - otherwise
    ; American "1,000" would be read as one-point-zero-zero-zero
    spacePattern := UseEuropeanFormat ? "^\d{1,3}(?: \d{3})*(?:,\d+)?$" : "^\d{1,3}(?: \d{3})*(?:\.\d+)?$"

    for line in StrSplit(text, "`n", "`r") {
        line := Trim(line)
        if (line = "")
            continue

        ; Check for numbers with spaces as thousand separators (e.g., "100 000")
        if RegExMatch(line, spacePattern) {
            ; Remove spaces and convert to number
            numStr := StrReplace(line, " ", "")
            if (UseEuropeanFormat)
                numStr := StrReplace(numStr, ",", ".")
            numbers.Push(Number(numStr))
            continue
        }

        ; Check for a whole line that is one number with other separators
        if (UseEuropeanFormat) {
            ; European format: 100.000,00 or 100.000
            if RegExMatch(line, "^\d{1,3}(?:\.\d{3})*(?:,\d+)?$") {
                numStr := StrReplace(line, ".", "")
                numStr := StrReplace(numStr, ",", ".")
                numbers.Push(Number(numStr))
                continue
            }
        } else {
            ; American format: 100,000.00 or 100,000
            if RegExMatch(line, "^\d{1,3}(?:,\d{3})*(?:\.\d+)?$") {
                numStr := StrReplace(line, ",", "")
                numbers.Push(Number(numStr))
                continue
            }
        }

        ; If no separators found, check for plain numbers
        if RegExMatch(line, "^\d+$") {
            numbers.Push(Number(line))
            continue
        }

        ; The line holds more than one value - pick the numbers out of it
        pos := 1
        while pos := RegExMatch(line, scanPattern, &match, pos) {
            numStr := match[0]
            if (UseEuropeanFormat) {
                numStr := StrReplace(numStr, ".", "")
                numStr := StrReplace(numStr, ",", ".")
            } else {
                numStr := StrReplace(numStr, ",", "")
            }
            numbers.Push(Number(numStr))
            pos += match.Len
        }
    }

    return numbers
}

; Improved function to extract plain numbers from text - better handling of decimals
ExtractNumbers(text) {
    numbers := []

    ; Use improved pattern based on current number format to better handle decimals
    pattern := UseEuropeanFormat ? "[-+]?\d+(?:,\d+)?" : "[-+]?\d+(?:\.\d+)?"

    pos := 1
    while pos := RegExMatch(text, pattern, &match, pos) {
        if pos > 0 {
            ; Convert to proper number based on format
            numberStr := match[0]

            ; Convert comma to period for Number() function if using European format
            if (UseEuropeanFormat && InStr(numberStr, ","))
                numberStr := StrReplace(numberStr, ",", ".")

            numbers.Push(Number(numberStr))
            pos += match.Len
        } else break
    }

    return numbers
}

; Improved format number function with consistent decimal places
FormatNumber(number) {
    ; Always round to 2 decimal places for consistency
    number := Round(number, 2)

    ; Determine if it's an integer after rounding
    isInteger := (number = Integer(number))

    ; Apply thousand separator formatting
    intPart := Integer(number)
    fracPart := number - intPart

    ; Format integer part with separators
    formattedInt := ""
    intStr := String(intPart)
    intLen := StrLen(intStr)

    loop parse, intStr {
        if (A_Index > 1 && Mod(intLen - A_Index + 1, 3) = 0)
            formattedInt .= UseEuropeanFormat ? "." : ","
        formattedInt .= A_LoopField
    }

    ; Even if it's an integer, we want to show 2 decimal places
    if (isInteger) {
        return formattedInt . (UseEuropeanFormat ? ",00" : ".00")
    } else {
        fracStr := SubStr(String(fracPart), 3)  ; Remove the leading "0." (adjusted for AHK v2)

        ; Ensure we always have exactly 2 decimal places
        if (StrLen(fracStr) > 2)
            fracStr := SubStr(fracStr, 1, 2)
        while (StrLen(fracStr) < 2)
            fracStr .= "0"

        return formattedInt . (UseEuropeanFormat ? "," : ".") . fracStr
    }
}

; Improved function to check if a string is a valid mathematical expression
IsCalculatableExpression(text) {
    ; Trim whitespace and check if empty
    if (text := Trim(text)) = ""
        return false

    ; Check if text contains math operators
    hasOperators := RegExMatch(text, "[\+\-\*\/\(\)\%\^]")

    ; If it has operators, it's likely an expression
    if hasOperators
        return true

    ; Improved pattern for detecting decimal numbers in either format
    decimalPattern := UseEuropeanFormat ? "^[-+]?\d+,\d+$" : "^[-+]?\d+\.\d+$"
    if RegExMatch(text, decimalPattern)
        return true

    ; If it's just a single number, it's technically an expression
    if RegExMatch(text, "^[-+]?\d+$")
        return true

    ; Otherwise, it's probably multiple numbers or invalid
    return false
}

; Expression calculation function with improved decimal handling
CalculateExpression(expression) {
    ; Clean up expression (trim whitespace)
    expression := Trim(expression)

    ; Handle European format if needed (replace comma with period for calculation)
    if (UseEuropeanFormat)
        expression := StrReplace(expression, ",", ".")

    try {
        ; Create temp file to evaluate
        tempFile := A_Temp "\calctemp.ahk"
        try FileDelete(tempFile)

        ; Create a script file to evaluate the expression
        FileAppend("#Requires AutoHotkey v2.0`nFileAppend(" . expression . ", '*')", tempFile)

        ; Run the script and capture output
        shell := ComObject("WScript.Shell")
        runCmd := '"' . A_AhkPath . '" /ErrorStdOut "' . tempFile . '"'

        result := Trim(shell.Exec(runCmd).StdOut.ReadAll())
        FileDelete(tempFile)

        ; Format according to current number format
        if (UseEuropeanFormat && InStr(result, ".")) {
            ; Handle decimal part properly
            result := StrReplace(result, ".", ",")
        }

        return result
    } catch as err {
        try FileDelete(tempFile)
        throw Error("Failed to evaluate: " . expression, -1, err.Message)
    }
}

; Function to show script's About information
ShowAboutDialog(*) {
    ; Read the script file
    scriptContent := FileRead(A_ScriptFullPath)

    ; Extract comment block (between /* and */)
    if RegExMatch(scriptContent, "\/\*[\s\S]*?\*\/", &commentBlock) {
        ; Clean up the comment block
        aboutText := RegExReplace(commentBlock[0], "\/\*|\*\/", "")  ; Remove /* and */
        aboutText := Trim(aboutText)  ; Remove extra whitespace

        ; Add current hotkey information
        aboutText .=  FormatHotkeyString(CurrentHotkey)

        ; Show the about box
        ShowMessage(aboutText, "About " . VarScriptName . " " . VarVersionNo)
    }
}

; Load saved settings from INI file
LoadSettings() {
    try {
        ; Load main hotkey
        savedHotkey := IniRead(ConfigFile, "Settings", "Hotkey", CurrentHotkey)
        if (savedHotkey != "")
            CurrentHotkey := savedHotkey

        ; Load number format setting
        savedFormat := IniRead(ConfigFile, "Settings", "UseEuropeanFormat", UseEuropeanFormat)
        if (savedFormat != "")
            UseEuropeanFormat := (savedFormat = "1" || savedFormat = "true") ? true : false
    } catch {
        ; If there's an error reading the file, stick with the defaults
    }
}

; Function to format a hotkey string for display
FormatHotkeyString(hotkeyString) {
    ; Walk the leading modifier symbols one at a time - chained StrReplace calls
    ; would turn the "+" in an already-inserted "Alt+" into "Shift+"
    modifierNames := Map("!", "Alt+", "^", "Ctrl+", "+", "Shift+", "#", "Win+")
    formattedString := ""
    pos := 1
    while (pos < StrLen(hotkeyString) && modifierNames.Has(SubStr(hotkeyString, pos, 1))) {
        formattedString .= modifierNames[SubStr(hotkeyString, pos, 1)]
        pos++
    }
    return formattedString . StrUpper(SubStr(hotkeyString, pos))
}

; Simplified function to let user change the hotkey
ChangeHotkey(*) {
    hotkeyGui := Gui("+AlwaysOnTop", VarScriptName " - Change Hotkey")
    hotkeyGui.SetFont("s10")
    hotkeyGui.Add("Text",, "Press the desired key combination:")

    calcHotkeyInput := hotkeyGui.Add("Hotkey", "vCalcHotkey w200", CurrentHotkey)

    okButton := hotkeyGui.Add("Button", "Default w100", "OK")
    okButton.OnEvent("Click", ProcessHotkeyChange)

    cancelButton := hotkeyGui.Add("Button", "x+5 w100", "Cancel")
    cancelButton.OnEvent("Click", (*) => (hotkeyGui.Destroy()))

    resetButton := hotkeyGui.Add("Button", "x10 y+10 w200", "Reset to Default (Alt+C)")
    resetButton.OnEvent("Click", (*) => (calcHotkeyInput.Value := "!c"))

    hotkeyGui.Show("AutoSize Center")

    ; Function to process the new hotkey
    ProcessHotkeyChange(*) {
        newHotkey := calcHotkeyInput.Value

        if (newHotkey = "") {
            ShowMessage("Please specify a hotkey combination.", "Error", 16)
            return
        }

        try {
            ; Disable old hotkey first
            try Hotkey CurrentHotkey, "Off"

            ; Enable new hotkey
            Hotkey newHotkey, CalculateSelectedTextOrShowCalculator

            ; Save the new hotkey to the INI file
            try IniWrite(newHotkey, ConfigFile, "Settings", "Hotkey")

            ; Update the current hotkey
            CurrentHotkey := newHotkey

            ; Update tray tip
            A_IconTip := VarScriptName " " VarVersionNo "`nHotkey: " FormatHotkeyString(CurrentHotkey)

            ; Format string for display
            displayHotkey := FormatHotkeyString(CurrentHotkey)

            ShowMessage("Hotkey changed to " displayHotkey)
            hotkeyGui.Destroy()
        } catch as err {
            ShowMessage("Invalid hotkey combination. Please try again.`nError: " err.Message, "Error", 16)
        }
    }
}

; Function to toggle between European and American number formats
ToggleNumberFormat(*) {
    ; Toggle the format setting
    global UseEuropeanFormat := !UseEuropeanFormat

    ; Save the setting to INI file
    try IniWrite(UseEuropeanFormat ? "1" : "0", ConfigFile, "Settings", "UseEuropeanFormat")

    ; Update tray menu item label
    FormatMenuLabel := UseEuropeanFormat ? "Switch to American Format (.) " : "Switch to European Format (,) "
    A_TrayMenu.Rename("3&", FormatMenuLabel)

    ; Update tray tooltip
    formatDescription := UseEuropeanFormat ? "European (,)" : "American (.)"
    A_IconTip := VarScriptName " " VarVersionNo "`nHotkey: " FormatHotkeyString(CurrentHotkey) "`nFormat: " formatDescription

    ; Show notification
    ShowMessage("Number format changed to " formatDescription)
}

; Simplified function to dismiss dialog boxes with Escape key
DismissMsgBox(*) {
    if WinActive("ahk_class #32770") {
        ; Send Escape key to dismiss the dialog
        Send "{Escape}"
    }
}

; Function to manage Windows startup
ManageStartup(add := true) {
    startupFolder := A_StartupCommon  ; Common startup folder for all users
    shortcutPath := startupFolder . "\" . VarScriptName . ".lnk"

    if (add) {
        try {
            if (A_IsCompiled) {
                ; Create the shortcut for an EXE
                FileCreateShortcut(A_ScriptFullPath, shortcutPath,
                                  A_WorkingDir,        ; Working directory
                                  "",                  ; No arguments
                                  "Launch " . VarScriptName . " on startup",
                                  A_ScriptDir . "\" . VarScriptName . ".ico")
            } else {
                ; Create the shortcut for an AHK script
                FileCreateShortcut(A_AhkPath, shortcutPath,
                                  A_WorkingDir,        ; Working directory
                                  '"' . A_ScriptFullPath . '"',  ; Arguments
                                  "Launch " . VarScriptName . " on startup",
                                  A_ScriptDir . "\" . VarScriptName . ".ico")
            }

            ; Update INI and menu
            try IniWrite("1", ConfigFile, "Settings", "AddedToStartup")

            ; Update menu
            UpdateStartupMenuItem(true)

            ShowMessage(VarScriptName . " has been added to startup.")
        } catch as err {
            ShowMessage("Unable to add " . VarScriptName . " to startup.`n`n"
                  . "Error details: " . err.Message)
        }
    } else {
        try {
            if FileExist(shortcutPath) {
                FileDelete(shortcutPath)

                try IniWrite("0", ConfigFile, "Settings", "AddedToStartup")

                ; Update menu
                UpdateStartupMenuItem(false)

                ShowMessage(VarScriptName . " has been removed from startup.")
            } else {
                ShowMessage(VarScriptName . " was not found in the startup folder.")
            }
        } catch as err {
            ShowMessage("Unable to remove " . VarScriptName . " from startup.`n`n"
                  . "Error details: " . err.Message)
        }
    }
}

; This function is no longer needed as we handle startup menu items directly in InitTrayMenu
; Keeping a stub for backward compatibility
UpdateStartupMenuItem(isInStartup) {
    ; Function intentionally left empty - startup menu items are now managed directly in InitTrayMenu
}

; Function to check if it's the first run and ask about startup
CheckFirstRunStartup() {
    try {
        askedAboutStartup := IniRead(ConfigFile, "Settings", "AskedAboutStartup", "0")
    } catch {
        askedAboutStartup := "0"  ; Default to not asked if can't read INI
    }

    if (askedAboutStartup = "0") {
        startupChoice := ShowMessage("Would you like CalculateInline to start automatically when Windows starts?`n`n"
                            . "This will allow you to use the " . FormatHotkeyString(CurrentHotkey)
                            . " shortcut anytime without having to manually start the program.",
                            VarScriptName " " VarVersionNo, "YesNo")

        try IniWrite("1", ConfigFile, "Settings", "AskedAboutStartup")

        if (startupChoice = "Yes")
            ManageStartup(true)
    }
}

; Initialize the tray menu
InitTrayMenu() {
    ; Try to set icon
    Try TraySetIcon(A_ScriptDir "\" VarScriptName ".ico")

    ; Clear the default menu completely
    A_TrayMenu.Delete()

    ; Create the custom menu from scratch exactly as requested
    A_TrayMenu.Add("Launch CalculateInline", ShowCalculator)
    A_TrayMenu.Add("Change Hotkey", ChangeHotkey)

    ; Add number format menu item
    FormatMenuLabel := UseEuropeanFormat ? "Switch to American Format (.) " : "Switch to European Format (,) "
    A_TrayMenu.Add(FormatMenuLabel, ToggleNumberFormat)

    A_TrayMenu.Add("About", ShowAboutDialog)

    ; Check if startup shortcut exists
    startupFolder := A_StartupCommon
    shortcutPath := startupFolder . "\" . VarScriptName . ".lnk"
    startupEnabled := FileExist(shortcutPath)

    ; Add startup item directly (not using the insert function)
    if (startupEnabled)
        A_TrayMenu.Add("Remove from Windows Startup", (*) => ManageStartup(false))
    else
        A_TrayMenu.Add("Add to Windows Startup", (*) => ManageStartup(true))

    ; Add exactly the items requested
    A_TrayMenu.Add("Reload Script", (*) => Reload())
    A_TrayMenu.Add("Edit Script", (*) => Run("notepad.exe " A_ScriptFullPath))
    A_TrayMenu.Add("Exit", (*) => ExitApp())

    ; Set default menu action
    A_TrayMenu.Default := "Launch CalculateInline"
    A_TrayMenu.ClickCount := 1  ; Single click to activate default

    ; Set initial tray tooltip
    formatDescription := UseEuropeanFormat ? "European (,)" : "American (.)"
    A_IconTip := VarScriptName " " VarVersionNo "`nHotkey: " FormatHotkeyString(CurrentHotkey) "`nFormat: " formatDescription
}

; ===== Script Initialization =====

; Load custom hotkey and other settings
LoadSettings()

; Register the initial hotkey
Hotkey CurrentHotkey, CalculateSelectedTextOrShowCalculator

; Register Escape key for message boxes
Hotkey "~Esc", DismissMsgBox, "On"

; Register the S/A/M/C/N/4 shortcuts for the results window
RegisterResultsShortcuts()

; Set up tray menu
InitTrayMenu()

; Show startup tooltip
ToolTip(VarScriptName " " VarVersionNo " loaded`nHotkey: " FormatHotkeyString(CurrentHotkey))
SetTimer () => ToolTip(), -3000

; Check if this is the first run to ask about startup
CheckFirstRunStartup()