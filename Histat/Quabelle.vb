Imports System.Collections
Imports Microsoft.Office.Interop.Excel

Public Class Quabelle
   Implements IDisposable

   Private xlMap As String
    Private xlApp As Microsoft.Office.Interop.Excel.Application
    ' Excel.ApplicationClass
    Private xlBook As Microsoft.Office.Interop.Excel.Workbook
    Private xlSheet As Microsoft.Office.Interop.Excel.Worksheet  'aktuelles Sheet
    Private rg As Microsoft.Office.Interop.Excel.Range
   Private rc As Int32     'rowCount
   Private cc As Int32     'columnCount
   Private hrc As Int32   'Anzahl Kopfzeilen 
   Private crc As Int32   'Anzahl Codezeilen 
   Private bTest As Boolean

   Private bDisposed As Boolean

   Public Sub New(ByVal xlmap As String)
      Try
         Me.xlMap = xlmap
            xlApp = New Microsoft.Office.Interop.Excel.ApplicationClass()
         xlBook = xlApp.Workbooks.Open(xlmap, , True)
         rc = countRows()
         cc = countColumns()
         bTest = testFormat1()        '  Setzen von hrc und crc
         xlSheet = xlBook.Worksheets(1)
         rg = xlSheet.Range("A1")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Public Sub setBook(ByVal xlmap As String)
      Try
         xlBook.Close(False, , )
         Me.xlMap = xlmap
         xlBook = xlApp.Workbooks.Open(xlmap, , True)
         rc = countRows()
         cc = countColumns()
         bTest = testFormat1()        '  Setzen von hrc und crc
         xlSheet = xlBook.Worksheets(1)
         rg = xlSheet.Range("A1")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Public ReadOnly Property RowCount() As Int32
      Get
         Return rc
      End Get
   End Property

   Public ReadOnly Property CodeRowCount() As Int32
      Get
         Return crc
      End Get
   End Property

   Public ReadOnly Property ColumnCount() As Int32
      Get
         Return cc
      End Get
   End Property

   Public ReadOnly Property Anmerkung() As Boolean
      Get
            Return (hrc - crc = 4)
      End Get
   End Property

   Public Function getItem(ByVal i As Int32, ByVal j As Int32) As String
      '  Gibt den Inhalt der Zelle in Zeile i und Spalte j zurück und setzt rg auf diese Zelle

      Dim shInd As Int32
      Dim cInd As Int32

      Try
         If i < 1 Or i > rc Or j < 1 Or j > cc Then
                MsgBox("getItem: Bereichsüberschreitung" & ControlChars.CrLf & "i = " & Convert.ToString(i) & ", j = " & _
                    Convert.ToString(j) & ControlChars.CrLf, MsgBoxStyle.Critical, "Histat-Import")
                HistatLog.WriteLine( _
                    Now & " getItem: Bereichsüberschreitung" & ControlChars.CrLf & "i = " & Convert.ToString(i) & ", j = " & Convert.ToString(j))
                Return Nothing
            Else
                If j = 1 Then
                    shInd = 1
                    cInd = 1
                Else
                    shInd = (j - 2) \ 255 + 1
                    cInd = j Mod 255
                    If cInd < 2 Then
                        cInd = cInd + 255
                    End If

                End If

                xlSheet = xlBook.Worksheets(shInd)
                rg = xlSheet.Cells(i, cInd)

                Return rg.Text
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getItemC(ByVal i As Int32, ByVal j As Int32, ByRef outComment As String) As String
        '  Gibt den Inhalt der Zelle in Zeile i und Spalte j zurück, speichert den Kommentar in outComment und setzt rg auf diese Zelle

        Dim shInd As Int32
        Dim cInd As Int32

        Try
            If i < 1 Or i > rc Or j < 1 Or j > cc Then
                MsgBox("getItem: Bereichsüberschreitung" & ControlChars.CrLf & "i = " & Convert.ToString(i) & ", j = " & _
                    Convert.ToString(j) & ControlChars.CrLf, MsgBoxStyle.Critical, "Histat-Import")
                HistatLog.WriteLine( _
                    Now & " getItem: Bereichsüberschreitung" & ControlChars.CrLf & "i = " & Convert.ToString(i) & ", j = " & Convert.ToString(j))
                Return Nothing
            Else
                If j = 1 Then
                    shInd = 1
                    cInd = 1
                Else
                    shInd = (j - 2) \ 255 + 1
                    cInd = j Mod 255
                    If cInd < 2 Then
                        cInd = cInd + 255
                    End If

                End If

                xlSheet = xlBook.Worksheets(shInd)
                rg = xlSheet.Cells(i, cInd)

                If (IsNothing(rg.Comment)) Then
                    outComment = Nothing
                Else
                    outComment = rg.Comment.Text
                End If
                Return rg.Text

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getNextRowEntry() As String
        'liefert den nächsten Eintrag in der aktuellen Zeile

        Dim str As String
        Dim shInd As Int32
        Dim rInd As Int32

        Try
            rInd = rg.Row
            shInd = xlSheet.Index
            If rg.Column < xlSheet.Columns.Count Then

                rg = rg.Range("B1")
                str = rg.Text
                Return str
            ElseIf shInd < xlBook.Worksheets.Count Then

                xlSheet = xlBook.Worksheets(shInd + 1)
                rg = xlSheet.Cells(rInd, 2)
                str = rg.Text
                Return str
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function getNextRowEntryC(ByRef outComment As Object) As String
        'liefert den nächsten Eintrag in der aktuellen Zeile, Kommentar in wird in outComment ausgegeben

        Dim str As String
        Dim shInd As Int32
        Dim rInd As Int32

        Try
            outComment = ""
            rInd = rg.Row
            shInd = xlSheet.Index
            If rg.Column < xlSheet.Columns.Count Then

                rg = rg.Range("B1")
                str = rg.Text

                If (IsNothing(rg.Comment)) Then
                    outComment = Nothing
                Else
                    outComment = rg.Comment.Text
                End If
                Return str
            ElseIf shInd < xlBook.Worksheets.Count Then

                xlSheet = xlBook.Worksheets(shInd + 1)
                rg = xlSheet.Cells(rInd, 2)
                str = rg.Text

                If (IsNothing(rg.Comment)) Then
                    outComment = Nothing
                Else
                    outComment = rg.Comment.Text
                End If
                Return str
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function getNextColEntry() As String
        'liefert den nächsten Eintrag in der aktuellen Spalte

        Try
            If rg.Row < rc Then
                rg = rg.Range("A2")
                Return rg.Text
            Else
                Return Nothing
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Public Function getNextRow() As String
        Dim str As String
        Dim rInd As Int32
        Try
            rInd = rg.Row + 1
            If rInd <= rc Then
                xlSheet = xlBook.Worksheets(1)
                rg = xlSheet.Cells(rInd, 1)
                str = rg.Text
                Return str
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function getDistinctValues(ByVal r As Int32) As Int32
        ' Anzahl verschiedener Werte in der Zeile r ab Spalte 2
        Dim al As Collections.ArrayList
        Dim str As String
        Dim cnt As Int32
        Try
            al = New ArrayList
            str = getItem(r, 2)
            Do While str <> Nothing
                If Not al.Contains(str) Then
                    al.Add(str)
                End If
                str = getNextRowEntry()
            Loop
            cnt = al.Count
            Return cnt

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function test() As Boolean
        Try
            If bTest Then
                bTest = bTest And Me.testHSName()
                bTest = bTest And Me.entriesCompl(hrc)
                test = bTest And Me.dataNum()
            Else
                test = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function entriesCompl(ByVal r As Int32) As Boolean
        ' Überprüft die Vollständigkeit der Zeilen 2 bis r
        Dim i As Int32
        Dim j As Int32
        Try
            entriesCompl = True
            For j = 1 To cc
                Debug.WriteLine(" Zeile 2, Spalte " & j & getItem(2, j))
                If Trim(getItem(2, j)) = "" Then
                    '                  MsgBox("Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                    '                      ControlChars.CrLf & "In Arbeitsblatt " & Me.xlSheet.Name & ", Zeile 2 und Spalte " & rg.Address.Substring(1, 1) & " fehlt ein Eintrag!", MsgBoxStyle.Critical, "Histat-Import")
                    HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                        ControlChars.CrLf & "In Arbeitsblatt " & Me.xlSheet.Name & ", Zeile 2 und Spalte " & rg.Column & " fehlt ein Eintrag!")
                    entriesCompl = False
                End If
                For i = 3 To r
                    rg = rg.Range("A2")
                    If Trim(rg.Text) = "" Then
                        '                      MsgBox("Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                        '                          ControlChars.CrLf & "In Arbeitsblatt " & Me.xlSheet.Name & ", Zeile " & rg.Row & " und Spalte " & rg.Address.Substring(1, 1) & " fehlt ein Eintrag!", MsgBoxStyle.Critical, "Histat-Import")
                        HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                            ControlChars.CrLf & "In Arbeitsblatt " & Me.xlSheet.Name & ", Zeile " & rg.Row & " und Spalte " & rg.Column & " fehlt ein Eintrag!")
                        entriesCompl = False
                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Private Function dataNum() As Boolean
        ' Überprüft die Zeitreiheneinträge auf numerisches Format
        Dim i As Int32
        Dim j As Int32
        Try
            dataNum = True
            For j = 2 To cc
                getItem(hrc + 1, j)
                For i = hrc + 1 To rc
                    If Not (IsNumeric(rg.Text) Or Trim(rg.Text) = "") Then
                        '                       MsgBox("Fehlerhafte Excel-Arbeitsmappe " & Me.xlMap & ":" & _
                        '                           ControlChars.CrLf & "Der Eintrag in Arbeitsblatt " & Me.xlSheet.Name & ", Zeile " & rg.Row & " und Spalte " & rg.Address.Substring(1, 1) & " ist fehlerhaft!", MsgBoxStyle.Critical, "Histat-Import")
                        HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                            ControlChars.CrLf & "Der Eintrag in Arbeitsblatt " & Me.xlSheet.Name & ", Zeile " & rg.Row & " und Spalte " & rg.Address.Substring(1, 1) & " ist fehlerhaft!")
                        dataNum = False
                    End If
                    rg = rg.Range("A2")
                Next
            Next

        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function countRows() As Int32
        Dim i As Int32
        Try
            xlSheet = xlBook.Worksheets(1)
            rg = xlSheet.Range("A1")
            i = 0
            Do While rg.Row() < xlSheet.Rows.Count
                If Trim(rg.Text) <> "" Then
                    rg = rg.Range("A2")
                    i += 1
                Else
                    Exit Do
                End If
            Loop
            Return i
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function countColumns() As Int32
        Dim ls As Int32
        Dim i As Int32
        Try
            With xlBook
                ls = .Worksheets.Count
                xlSheet = .Worksheets(ls)
            End With
            If xlSheet.Range("IV2").Text <> "" Then
                Return ls * (xlSheet.Columns.Count - 1) + 1
            Else
                rg = xlSheet.Range("B2")
                i = 0
                Do While Trim(rg.Text) <> "" And rg.Column < xlSheet.Columns.Count
                    rg = rg.Range("B1")
                    i += 1
                Loop
                Return (ls - 1) * (xlSheet.Columns.Count - 1) + 1 + i
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function testHSName() As Boolean
        Dim HSName As String
        Dim i As Int32
        Try
            testHSName = True
            With xlBook
                xlSheet = .Worksheets(1)
                HSName = xlSheet.Range("A1").Text
                For i = 2 To .Worksheets.Count
                    xlSheet = .Worksheets(i)
                    If xlSheet.Range("A1").Text <> HSName Then
                        '   MsgBox("Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                        '   ControlChars.CrLf & "Hauptschlüsselname auf Seite " & i & _
                        '    " stimmt nicht mit dem Hauptschlüsselname auf Seite 1 überein!", MsgBoxStyle.Critical, "Histat-Import")
                        HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                            ControlChars.CrLf & "Hauptschlüsselname auf Seite " & i & _
                            " stimmt nicht mit dem Hauptschlüsselname auf Seite 1 überein!")
                        testHSName = False
                    End If
                Next
            End With
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Private Function testFormat1() As Boolean
        Dim i As Int32

        Try
            testFormat1 = True
            xlSheet = xlBook.Worksheets(1)
            rg = xlSheet.Range("A2")
            hrc = 0

            For i = 2 To rc
                If rg.Text = "Quelle" Then
                    crc = i - 2
                    hrc = i + 1
                    Exit For
                Else
                    rg = rg.Range("A2")
                End If
            Next
            If hrc = 0 Then
                '    MsgBox("Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                '    ControlChars.CrLf & "Quelle nicht vorhanden!", MsgBoxStyle.Critical, "Histat-Import")
                HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                    ControlChars.CrLf & "Eintrag Quelle nicht vorhanden!")
                hrc = -1
                crc = -1
                testFormat1 = False

            Else
                rg = rg.Range("A2")
                If rg.Text = "Anmerkung" Then
                    hrc = hrc + 1
                    rg = rg.Range("A2")
                End If

                If rg.Text <> "Tabelle" Then
                    HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                        ControlChars.CrLf & "Eintrag Tabelle nicht vorhanden!")
                    hrc = -1
                    crc = -1
                    testFormat1 = False
                End If

                If hrc = rc Then
                    '    MsgBox("Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                    '       ControlChars.CrLf & "Es sind keine Datenzeilen vorhanden!", MsgBoxStyle.Critical, "Histat-Import")
                    HistatLog.WriteLine(Now & " Excel-Arbeitsmappe " & Me.xlMap & " ist fehlerhaft!" & _
                        ControlChars.CrLf & "Es sind keine Datenzeilen vorhanden!")
                    testFormat1 = False
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function


   Public Sub Dispose() Implements IDisposable.Dispose
      Try
         If Not bDisposed Then
            Finalize()
            System.GC.SuppressFinalize(Me)
         End If
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Protected Overrides Sub Finalize()
      Try
         rg = Nothing
         xlSheet = Nothing
         If Not xlBook Is Nothing Then
            xlBook.Close(False)
            xlBook = Nothing
         End If
         If Not xlApp Is Nothing Then
            xlApp.Quit()
            xlApp = Nothing
         End If
         bDisposed = True
         MyBase.Finalize()
      Catch ex As Exception
         Try
            xlApp.Quit()
            xlApp = Nothing
            bDisposed = True
            MyBase.Finalize()
         Catch ex1 As Exception
            Throw ex
         End Try
      End Try

   End Sub
End Class
