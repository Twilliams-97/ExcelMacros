'Place at the start of the macro
Public calcMode As Long
Public pageBreakStatus As Boolean

Sub SettingsStartOfMacro()

With Application

    calcMode = .Calculation
    pageBreakStatus = ActiveSheet.DisplayPageBreaks

    'Turn calculation mode to manual
    .Calculation = xlCalculationManual

    'Turn off screen updating (i.e. no annoying screen flash)
    .ScreenUpdating = False

    'Alert windows will not be displayed
    .DisplayAlerts = False

End With

'Turn off page breaks on active sheet
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub SettingsEndOfMacro()

With Application

    'Return calculation to automatic
    .Calculation = calcMode

    'Turn screen updating back on
    .ScreenUpdating = True

    'Enable alerts to be shown
    .DisplayAlerts = True

End With

'Reset page breaks on active sheet
ActiveSheet.DisplayPageBreaks = pageBreakStatus

End Sub
