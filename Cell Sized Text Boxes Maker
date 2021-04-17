'Created by Tom Williams

Sub CellsToTextBoxes()

    Dim c As Range

    For Each c In Selection

        With c

            ActiveSheet.Shapes.AddTextbox _
            msoTextOrientationHorizontal, .Left, _
            .Top, .Width, .Height

        End With

    Next c

End Sub

