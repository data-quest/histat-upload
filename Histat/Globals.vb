Imports System.Text
Imports System.Data
Imports System.IO
Imports System.Security.Cryptography

Module Globals

    Enum tStatus As Byte

        onDelete
        onImport
        onDataImport
        onMetadataUpdate
    End Enum

    Public status As tStatus
    Public Const Schluessellaenge As Byte = 32
    Public Const conStr As String = "DSN=histat"

    Public sf As StartForm
    Public anm As Anmeldung
    Public md1 As Metadaten1
    Public md2 As Metadaten2
    Public md3 As Metadaten3

    Public da As Dateiauswahl
    Public pa As Projektauswahl

    Public HistatLog As StreamWriter
    Public onAbbr As Boolean = False

    Public xlsArr() As String     'Pfade der zu importierenden Excel-Dateien
    Public ID_Thema As Int32
    Public ID_Projekt As String
    Public Projektautor As String
    Public Projektname As String
    Public Projektbeschreibung As String
    Public ID_Zeit As Integer
    Public Veroeff As String
    Public Publikationsjahr As String
    Public Untersuch As String
    Public Quellen As String
    Public Untergliederung As String
    Public ZAStd As String
    Public DatA As String
    Public DatB As String
    Public Bearb As String
    Public Bem As String
    Public Zugang As String
    Public Fundort As String
    Public Anmerkungsteil As String
    Public Exportierbar As Byte
    Public DateiPfad As String
    Public DateiName As String

    Private pr As Diagnostics.Process

    Public Function checkXlTabs(ByRef anz_Zr As Int32, ByRef zeitraum As String) As Boolean
        ' Überprüft, ob Dateien angegeben wurden und ggfls. deren Korrektheit.
        '
        Dim i As Int32
        Dim l As Int32
        Dim ub As Int32
        Dim quab As Quabelle
        Dim xlFile As String
        Dim azr As Int32
        Dim von As String
        Dim bis As String

        Try
            ub = xlsArr.GetUpperBound(0)
            If ub < 0 Then
                Return False
            End If
            checkXlTabs = True
            da.SPanel2.Text = "checkXlTabs"
            xlFile = xlsArr(0)
            l = InStrRev(xlFile, "\")
            da.SPanel1.Text = Mid(xlFile, l + 1)
            da.SBar.Refresh()

            quab = New Quabelle(xlsArr(0))
            With quab
                bis = ""
                von = ""
                If Not .test() Then
                    checkXlTabs = False
                Else
                    azr = .ColumnCount - 1
                    If .Anmerkung Then
                        von = .getItem(.CodeRowCount + 5, 1)
                    Else
                        von = .getItem(.CodeRowCount + 4, 1)
                    End If
                    bis = .getItem(.RowCount, 1)
                End If

                For i = 1 To ub
                    xlFile = xlsArr(i)
                    l = InStrRev(xlFile, "\")
                    da.SPanel1.Text = Mid(xlFile, l + 1)
                    da.SBar.Refresh()

                    .setBook(xlsArr(i))
                    If Not .test() Then
                        checkXlTabs = False
                    Else
                        azr = azr + .ColumnCount - 1
                        If .Anmerkung Then
                            von = min(von, .getItem(.CodeRowCount + 5, 1))
                        Else
                            von = min(von, .getItem(.CodeRowCount + 4, 1))
                        End If
                        bis = max(bis, .getItem(.RowCount, 1))
                    End If
                Next
                .Dispose()
            End With
            anz_Zr = azr
            zeitraum = Left(von, 4) & " - " & Left(bis, 4)
            If Not checkXlTabs Then
                HistatLog.WriteLine()
            End If
        Catch ex As Exception
            If Not quab Is Nothing Then
                quab.Dispose()
                quab = Nothing
            End If
            Throw ex
        End Try
    End Function


    Public Sub closeForms()
        Try
            onAbbr = True
            If Not IsNothing(anm) Then
                anm.Close()
            End If
            If Not IsNothing(md1) Then
                md1.Close()
            End If
            If Not IsNothing(md2) Then
                md2.Close()
            End If
            If Not IsNothing(md3) Then
                md3.Close()
            End If
            If Not IsNothing(da) Then
                da.Close()
            End If
            If Not IsNothing(pa) Then
                pa.Close()
            End If
            '            HistatLog.WriteLine()
            HistatLog.Flush()
            sf.Show()
        Catch ex As Exception
            Throw ex
        End Try

    End Sub


    Public Function CreateUniqueKey(ByVal md5 As MD5CryptoServiceProvider) As String
        Dim hResult As Byte() = md5.ComputeHash(Encoding.ASCII.GetBytes(DateTime.Now.Ticks.ToString))
        Dim i As Integer
        Dim uniqueKey As String = ""
        Try
            For i = 0 To hResult.GetUpperBound(0)
                uniqueKey = String.Concat(uniqueKey, hex_0(hResult(i)))
            Next
            Console.WriteLine()
            Return uniqueKey
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    Private Function hex_0(ByVal b As Byte) As String
        Try
            Dim strHex As String
            strHex = (Hex(b)).PadLeft(2, "0")
            Return strHex
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function min(ByVal str1 As String, ByVal str2 As String) As String
        Try
            If String.Compare(str1, str2) < 0 Then
                Return str1
            Else
                Return str2
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    Private Function max(ByVal str1 As String, ByVal str2 As String) As String
        Try
            If String.Compare(str1, str2) > 0 Then
                Return str1
            Else
                Return str2
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Public Function GetDatei(ByVal filePath As Object) As Byte()
        Dim fs As FileStream
        Dim br As BinaryReader

        Dim datei() As Byte

        If filePath Is Nothing Then
            datei = Nothing
        Else
            fs = New FileStream(filePath, FileMode.Open, FileAccess.Read)
            br = New BinaryReader(fs)

            datei = br.ReadBytes(fs.Length)

            br.Close()
            fs.Close()

        End If
        Return datei
    End Function

    Public Sub openLog()

        If Not pr Is Nothing Then
            Try
                pr.CloseMainWindow()
                'pr.Kill()
            Catch ex As Exception
            End Try
        End If
        pr = Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath & "\Histat.log")

    End Sub

End Module

