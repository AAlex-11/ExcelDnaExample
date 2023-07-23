Imports ExcelDna.Integration
Imports ExcelDna.ComInterop.ComServer
Imports Microsoft.Office.Interop.Excel

Public Module Module1

    <ExcelFunction(Category:="Colorize", IsMacroType:=True)>
    <ExcelCommand(ShortCut:="^D", MenuText:="Colorize", Name:="Colorize")>
    Public Sub Colorize(Num As Integer)
        Try
            Debug.WriteLine("Start")
            Dim Xl As Application = ExcelDnaUtil.Application
            Dim Range1 As Range = Xl.Range($"A1:Z{Num}")
            Dim i As Integer = 0
            Dim RND As New Random(255)
            For Each One As Range In Range1
                Debug.WriteLine(One.Value)
                One.Interior.Color = RGB(RND.Next Mod 255, RND.Next Mod 255, RND.Next Mod 255)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Module
