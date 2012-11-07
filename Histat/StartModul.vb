Imports System
Imports System.IO
Imports System.Windows.Forms

Module StartModul
    Sub Main()

        Try
            HistatLog = New StreamWriter(Application.StartupPath & "\Histat.log")
            Debug.Print(Application.StartupPath)
            HistatLog.WriteLine(Now & " HistatLog")
            HistatLog.WriteLine()
        Catch ex As SystemException
            MsgBox("Es ist keine Zugriffsmöglickeit auf die Log-Datei " & Application.StartupPath & "\Histat.log" & " gegeben." & vbCrLf & "Bitte überprüfen Sie die Sicherheitseinstellungen in den Eigenschaften der Datei.", MsgBoxStyle.Critical, "Histat")
            Exit Sub
        End Try

        Try
            sf = New StartForm()
            sf.Show()
            Application.Run()
        Catch ex As OdbcException
            HistatLog.WriteLine(Now & (ex).GetType.Name & ControlChars.CrLf & "Der Datenquellenname wurde nicht gefunden!" & _
                ControlChars.CrLf & "Bitte erstellen Sie erst eine Odbc-Datenquelle!")
            MsgBox((ex).GetType.Name & ControlChars.CrLf & "Der Datenquellenname wurde nicht gefunden!" & _
                ControlChars.CrLf & "Bitte erstellen Sie erst eine Odbc-Datenquelle!", _
            MsgBoxStyle.Critical, "Histat")
            closeForms()
        Catch ex As Exception
            HistatLog.WriteLine(Now & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
            closeForms()
        End Try
    End Sub

End Module
