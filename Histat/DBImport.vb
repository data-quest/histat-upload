Imports System.Data.Odbc

Imports System.Security.Cryptography
Imports System.Math


Module DBImport
   Private schlArr() As String
   Private ID_Project As String
   Private md5 As MD5CryptoServiceProvider
   Private quab As Quabelle
   Private conn As OdbcConnection
   Private ds As DataSet

   Private adapt As OdbcDataAdapter

   Public Function ImportProject(ByVal azr As String, ByVal zeitraum As String) As String
      Try
         md5 = New MD5CryptoServiceProvider()
         conn = New OdbcConnection(conStr)
         conn.Open()
         ds = New DataSet()
         CreateDS(conn)
         ImportProject = Insert_Aka_Projekte(azr, zeitraum)
         conn.Close()
      Catch ex As Exception
         ImportProject = Nothing
         Throw ex
      End Try
   End Function

   Public Function ImportData(ByVal strXl As String, ByRef exc As Exception, ByVal strID_Project As String, ByVal anzZR As Int32, ByVal zr As String) As Boolean
        Dim tbl As System.Data.DataTable

      Try
         ID_Project = strID_Project
         If IsNothing(conn) Then             ' ImportProjekt wurde nicht aufgerufen!
            md5 = New MD5CryptoServiceProvider()
            conn = New OdbcConnection(conStr)
            conn.Open()
         End If

         If IsNothing(ds) Then
            ds = New DataSet()
            CreateDS(conn)
         Else
            For Each tbl In ds.Tables
               tbl.Clear()
            Next
         End If

         quab = New Quabelle(strXl)
         CreateEntries()
         quab.Dispose()
         quab = Nothing

         InsertDB(anzZR, zr)
         conn.Close()
         ImportData = True
      Catch ex As Exception
         ImportData = False
         exc = ex
         If Not quab Is Nothing Then
            quab.Dispose()
            quab = Nothing
         End If
      End Try
   End Function

   Public Sub dispose()
      Try
         md5 = Nothing
         ds = Nothing
         quab = Nothing
         conn = Nothing
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Private Sub CreateDS(ByVal con As OdbcConnection)
      Dim sSql As String
        Dim dt As System.Data.DataTable
      Dim uc As UniqueConstraint

      Try
         sSql = "SELECT * FROM Aka_Projekte WHERE ID_Projekt IS NULL;"
         adapt = New OdbcDataAdapter(sSql, con)
         adapt.Fill(ds, "DS_Aka_Projekte")

         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_Schluesselmaske WHERE ID_HS IS NULL;"
         adapt.Fill(ds, "DS_Aka_Schluesselmaske")

         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_Codes WHERE ID_CodeKuerz IS NULL;"
         adapt.Fill(ds, "DS_Aka_Codes")

         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_SchluesselCode WHERE ID_HS IS NULL;"
         adapt.Fill(ds, "DS_Aka_SchluesselCode")

         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_CodeInhalt WHERE ID_CodeKuerz IS NULL;"
         adapt.Fill(ds, "DS_Aka_CodeInhalt")

         adapt.SelectCommand.CommandText = "SELECT * FROM Lit_ZR WHERE ID_HS IS NULL;"
         adapt.Fill(ds, "DS_Lit_ZR")
         dt = ds.Tables("DS_Lit_ZR")
         uc = New UniqueConstraint("Keys_Lit_ZR", _
               New DataColumn() {dt.Columns("ID_HS"), dt.Columns("Schluessel")})
         dt.Constraints.Add(uc)

         adapt.SelectCommand.CommandText = "SELECT * FROM Daten__Aka WHERE ID_HS IS NULL;"
         adapt.Fill(ds, "DS_Daten__Aka")
         dt = ds.Tables("DS_Daten__Aka")
         uc = New UniqueConstraint("Keys_Daten__Aka", _
               New DataColumn() {dt.Columns("ID_HS"), dt.Columns("Schluessel"), dt.Columns("Jahr_Sem")})
         dt.Constraints.Add(uc)

      Catch ex As Exception
         Throw ex
      End Try

   End Sub


   Private Sub CreateEntries()
        Dim Pos As Int32
        Dim i As Int32
        Dim j As Int32
        Dim Zchn As Int32
        Dim ID_CodeKuerz As String
        Dim ID_HS As String
        Dim ICode As Int32
        Dim Code As String
        Dim cc As Int32
        Dim CodeBezeichnung As String
        Dim dt As System.Data.DataTable
        Dim Schluessel As String
        Dim iQuelle As Int32
        Dim bAnm As Boolean
        Dim startDaten As Int32
        Dim endDaten As Int32
        Dim dData As Double
        Dim strData As String
        Dim strJS As String
        Dim strComment As String

        ID_CodeKuerz = ""
        ID_HS = ""
        CodeBezeichnung = ""
        Code = ""
        Schluessel = ""
        strData = ""
        strJS = ""

        Try
            ID_HS = CreateUniqueKey(md5)
            Create_Aka_Schluesselmaske(ID_HS)
            cc = quab.ColumnCount
            ReDim schlArr(cc)    'Inhalte löschen
            For j = 2 To cc
                schlArr(j) = ""
            Next
            Pos = 1
            dt = ds.Tables("DS_Aka_CodeInhalt")
            For i = 2 To quab.CodeRowCount + 1
                Zchn = Floor(Log(quab.getDistinctValues(i), 10)) + 1

                da.SPanel2.Text = "Create_Aka_Codes"
                da.Refresh()

                ID_CodeKuerz = ""
                Create_Aka_Codes(i, Zchn, ID_CodeKuerz)
                Create_Aka_SchluesselCode(ID_HS, ID_CodeKuerz, Pos)

                Pos = Pos + Zchn
                ICode = 0
                For j = 2 To cc
                    da.SPanel2.Text = "Create_Aka_CodeInhalt"
                    da.Refresh()

                    Code = ""
                    CodeBezeichnung = quab.getItem(i, j)
                    If Not ExistsCodeBezeichnung(CodeBezeichnung, ID_CodeKuerz, Code, dt) Then
                        Code = Convert.ToString(ICode).PadLeft(Zchn, "0")
                        Create_Aka_CodeInhalt(ID_CodeKuerz, Code, CodeBezeichnung, Zchn)
                        ICode += 1
                    End If
                    schlArr(j) = schlArr(j) & Code
                Next
            Next

            iQuelle = quab.CodeRowCount + 2
            bAnm = quab.Anmerkung

            da.SPanel2.Text = "Create_Lit_ZR"
            da.Refresh()

            For j = 2 To cc
                Schluessel = schlArr(j)
                If Schluessel.Length > Schluessellaenge Then
                    Throw New Exception("Schlüsselraumüberschreitung!")
                End If
                Schluessel = Schluessel.PadRight(Schluessellaenge, "0")
                schlArr(j) = Schluessel

                If bAnm Then
                    Create_Lit_ZR_Anm(j, ID_HS, Schluessel, iQuelle)
                Else
                    Create_Lit_Zr(j, ID_HS, Schluessel, iQuelle)
                End If

            Next

            startDaten = IIf(bAnm, iQuelle + 3, iQuelle + 2)
            endDaten = quab.RowCount

            For i = startDaten To endDaten
                da.SPanel2.Text = "Create_Daten__Aka: Zeile " & Convert.ToString(i)
                da.Refresh()
                strJS = quab.getItem(i, 1)
                For j = 2 To cc
                    strComment = ""
                    strData = quab.getNextRowEntryC(strComment)
                    If String.Compare(Trim(strData), "", True) <> 0 Then
                        dData = Convert.ToDouble(strData)
                        Create_Daten__Aka(ID_HS, schlArr(j), strJS, strData, strComment)
                    End If
                Next
            Next
        Catch ex As Exception
            MsgBox("DBImport.CreateEntries: " & _
                "i = " & i & vbCrLf & _
                "j = " & j & vbCrLf & _
                "ID_CodeKuerz = " & ID_CodeKuerz & vbCrLf & _
                "ID_HS = " & ID_HS & vbCrLf & _
                "Code = " & Code & vbCrLf & _
                "cc = " & cc & vbCrLf & _
                "CodeBezeichnung = " & CodeBezeichnung & vbCrLf & _
                "Schluessel = " & Schluessel & vbCrLf & _
                "strData = " & strData & vbCrLf & _
                "strJS = " & strJS)
            HistatLog.WriteLine(Now & "DBImport.CreateEntries: " & _
                "i = " & i & vbCrLf & _
                "j = " & j & vbCrLf & _
                "ID_CodeKuerz = " & ID_CodeKuerz & vbCrLf & _
                "ID_HS = " & ID_HS & vbCrLf & _
                "Code = " & Code & vbCrLf & _
                "cc = " & cc & vbCrLf & _
                "CodeBezeichnung = " & CodeBezeichnung & vbCrLf & _
                "Schluessel = " & Schluessel & vbCrLf & _
                "strData = " & strData & vbCrLf & _
                "strJS = " & strJS & vbCrLf)
            Throw ex
        End Try
   End Sub


    Private Sub Create_Aka_Schluesselmaske(ByVal id_hs As String)
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try
            dt = ds.Tables("DS_Aka_Schluesselmaske")
            dr = dt.NewRow()
            dr("ID_HS") = id_hs
            dr("Name") = quab.getItem(1, 1)
            dr("ID_Projekt") = ID_Project

            dt.Rows.Add(dr)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Create_Aka_Codes(ByVal i As Integer, ByVal zchn As Int32, ByRef id_codekuerz As String)
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try
            dt = ds.Tables("DS_Aka_Codes")
            dr = dt.NewRow()
            id_codekuerz = CreateUniqueKey(md5)
            dr("ID_CodeKuerz") = id_codekuerz
            dr("Codebeschreibung") = quab.getItem(i, 1)
            dr("Zeichen") = zchn
            dt.Rows.Add(dr)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub Create_Aka_SchluesselCode(ByVal id_hs As String, ByVal id_codekuerz As String, ByVal pos As Int32)
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try
            dt = ds.Tables("DS_Aka_SchluesselCode")
            dr = dt.NewRow()
            dr("ID_HS") = id_hs
            dr("ID_CodeKuerz") = id_codekuerz
            dr("Position") = pos
            dt.Rows.Add(dr)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Create_Aka_CodeInhalt(ByVal id_codekuerz As String, ByVal code As String, ByVal codebez As String, ByVal zchn As String)
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try
            dt = ds.Tables("DS_Aka_CodeInhalt")
            dr = dt.NewRow()
            dr("ID_CodeKuerz") = id_codekuerz
            dr("Code") = code
            dr("CodeBezeichnung") = codebez
            dt.Rows.Add(dr)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function ExistsCodeBezeichnung(ByVal codeBez As String, ByVal id_codekuerz As String, ByRef code As String, ByVal dt As System.Data.DataTable) As Boolean
        ' Überprüft, ob die Codebezeichnung schon existiert. 
        ' In diesem Fall wird code gesetzt und true zurückgegeben.

        Dim dr As DataRow

        Try
            For Each dr In dt.Rows
                If String.Compare(dr("Id_CodeKuerz"), id_codekuerz, False) = 0 And _
                   String.Compare(dr("Codebezeichnung"), codeBez, True) = 0 Then
                    code = dr("Code")
                    Return True
                    Exit For
                End If
            Next
            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Sub Create_Lit_Zr(ByVal c As Int32, ByVal id_hs As String, ByVal schluessel As String, ByVal iQuelle As Int32)
        Dim dt As System.Data.DataTable
        Dim dr As DataRow
        Dim Quelle As String
        Dim Tabelle As String

        Try
            Quelle = quab.getItem(iQuelle, c)
            Tabelle = quab.getItem(iQuelle + 1, c)
            dt = ds.Tables("DS_Lit_ZR")
            dr = dt.NewRow()
            dr("ID_HS") = id_hs
            dr("Schluessel") = schluessel
            dr("Quelle") = Quelle
            dr("Tabelle") = Tabelle
            dt.Rows.Add(dr)
        Catch ex As Exception
            If ex.Message.EndsWith("bereits vorhanden.") Then
                Throw New Exception("Create_Lit_Zr: Ein Eintrag mit dem Schlüssel '" & schluessel & "' ist bereits vorhanden." & vbCrLf & _
                   "Der doppelte Eintrag befindet sich auf Tabellenblatt " & (c - 1) \ 255 + 1 & _
                   " an Spalte " & (c - 1) Mod 255 + 1 & ".")
            Else
                Throw ex
            End If
        End Try
    End Sub

    Private Sub Create_Lit_ZR_Anm(ByVal c As Int32, ByVal id_hs As String, ByVal schluessel As String, ByVal iQuelle As Int32)
        Dim Quelle As String
        Dim Anmerkung As String
        Dim Tabelle As String
        Dim dt As System.Data.DataTable
        Dim dr As DataRow

        Try
            Quelle = quab.getItem(iQuelle, c)
            Anmerkung = quab.getItem(iQuelle + 1, c)
            Tabelle = quab.getItem(iQuelle + 2, c)
            dt = ds.Tables("DS_Lit_ZR")
            dr = dt.NewRow()
            dr("ID_HS") = id_hs
            dr("Schluessel") = schluessel
            dr("Quelle") = Quelle
            dr("Anmerkung") = Anmerkung
            dr("Tabelle") = Tabelle
            dt.Rows.Add(dr)
            Debug.WriteLine(dt.ToString)
        Catch ex As Exception
            If ex.Message.EndsWith("bereits vorhanden.") Then
                Throw New Exception("Create_Lit_Zr: Ein Eintrag mit dem Schlüssel '" & schluessel & "' ist bereits vorhanden." & vbCrLf & _
                   "Der doppelte Eintrag befindet sich auf Tabellenblatt " & (c - 1) \ 255 + 1 & _
                   " an Spalte " & (c - 1) Mod 255 + 1 & ".")
            Else
                Throw ex
            End If
        End Try
    End Sub

   Private Sub Create_Daten__Aka(ByVal id_hs As String, ByVal schluessel As String, ByVal str_jahr_sem As String, ByVal dData As Double, ByRef strAnmerkung As Object)
        Dim dt As System.Data.DataTable
      Dim dr As DataRow

      Try
         dt = ds.Tables("DS_Daten__Aka")
         dr = dt.NewRow()
         dr("ID_HS") = id_hs
         dr("Schluessel") = schluessel
         dr("Jahr_Sem") = checkJahrSem(str_jahr_sem)
         dr("Data") = dData
         dr("Anmerkung") = strAnmerkung
         dt.Rows.Add(dr)
      Catch ex As Exception
         Throw ex
      End Try

   End Sub


   Private Function checkJahrSem(ByVal str As String) As String
        str = Trim(UCase(str))
      Try
         If str Like "####" Or str Like "####[1-3]" Then
            Return str
         ElseIf str Like "#### SS" Then
            Return Left(str, 4) & "1"
         ElseIf str Like "#### WS" Then
            Return Left(str, 4) & "2"
         ElseIf str Like "#### ZS" Then
            Return Left(str, 4) & "3"
         ElseIf str Like "####[/]I" Or str Like "####[/]II" Or str Like "####[/]III" Or str Like "####[/]IV" Then
            Return str
         ElseIf str Like "####[/]##" And Mid(str, 6, 2) >= "01" And Mid(str, 6, 2) <= "12" Then
            Return str
         ElseIf str Like "####[/]KW##" And Mid(str, 8, 2) >= "01" And Mid(str, 8, 2) <= "53" Then
            Return str
         ElseIf str Like "####[/]##[/]##" And Mid(str, 6, 2) >= "01" And Mid(str, 6, 2) <= "12" _
                     And Mid(str, 9, 2) >= "01" And Mid(str, 9, 2) <= "31" Then
            Return str
         ElseIf str Like "####[-]####" And Mid(str, 1, 4) <= Mid(str, 6, 4) Then
            Return str
         Else
            Throw New Exception("checkJahrSem: Jahr_Sem = " & str)
            '            Return Nothing
         End If

      Catch ex As Exception
         Throw New Exception("checkJahrSem: Jahr_Sem = " & str)
      End Try
   End Function

   Private Sub InsertDB(ByVal anzZr As Int32, ByVal zr As String)
      Dim cmd As OdbcCommand
      Try
         Insert_Aka_Schluesselmaske()
         Insert_Aka_Codes()
         Insert_Aka_SchluesselCode()
         Insert_Aka_CodeInhalt()
         Insert_Lit_ZR()
         Insert_Daten__Aka()
         If status = Globals.tStatus.onDataImport Then
            cmd = New OdbcCommand( _
            "UPDATE Aka_Projekte SET Anzahl_Zeitreihen = " & anzZr & _
            ", Zeitraum = '" & zr & _
            "' WHERE ID_Projekt = '" & ID_Projekt & "';", conn)
            cmd.ExecuteNonQuery()
         End If
      Catch ex As Exception
         Throw ex
      End Try
   End Sub



   Private Function Insert_Aka_Projekte(ByVal azr As String, ByVal zeitraum As String) As String
        '   liefert bei erfolgreichem Eintrag des Projekts  ID_Project zurück, null sonst.
        Dim cmd As OdbcCommand
        Dim param As OdbcParameter
        Dim DateiInhalt() As Byte
        Dim i As Int16
        Dim msg As String

        Try
            da.SPanel2.Text = "Insert_Aka_Projekte"
            da.Refresh()
            ID_Project = CreateUniqueKey(md5)
            If DateiName Is Nothing Then
                cmd = New OdbcCommand("Insert Into Aka_Projekte (" & _
                   "ID_Projekt, ID_Thema, ID_Zeit, Projektautor, Projektname, Projektbeschreibung, Veroeffentlichung, " & _
                   "Untersuchungsgebiet, Quellen, Untergliederung, ZA_Studiennummer, Datum_der_Archivierung, Datum_der_Bearbeitung, " & _
                   "Publikationsjahr, Bearbeiter_im_ZA, Bemerkungen, Zugangsklasse, Anzahl_Zeitreihen, Zeitraum, exportable, Fundort, " & _
                   "Anmerkungsteil) Values ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", conn)
            Else
                cmd = New OdbcCommand("Insert Into Aka_Projekte (" & _
                   "ID_Projekt, ID_Thema, ID_Zeit, Projektautor, Projektname, Projektbeschreibung, Veroeffentlichung, " & _
                   "Untersuchungsgebiet, Quellen, Untergliederung, ZA_Studiennummer, Datum_der_Archivierung, Datum_der_Bearbeitung, " & _
                   "Publikationsjahr, Bearbeiter_im_ZA, Bemerkungen, Zugangsklasse, Anzahl_Zeitreihen, Zeitraum, exportable, Fundort, " & _
                   "Anmerkungsteil, datei_inhalt, datei_name) Values ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);", conn)
            End If
            '
            param = New OdbcParameter("@ID_Projekt", OdbcType.VarChar, 32)
            param.Value = ID_Project
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@ID_Thema", OdbcType.Int, 11)
            param.Value = ID_Thema
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@ID_Zeit", OdbcType.Int, 11)
            param.Value = ID_Zeit
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Projektautor", OdbcType.VarChar, 255)
            param.Value = Projektautor
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Projektname", OdbcType.VarChar, 255)
            param.Value = Projektname
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Projektbeschreibung", OdbcType.Text)
            param.Value = Projektbeschreibung
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Veroeffentlichung", OdbcType.Text)
            param.Value = Veroeff
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Untersuchungsgebiet", OdbcType.Text)
            param.Value = Untersuch
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Quellen", OdbcType.Text)
            param.Value = Quellen
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Untergliederung", OdbcType.Text)
            param.Value = Untergliederung
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@ZA_Studiennummer", OdbcType.VarChar, 32)
            param.Value = ZAStd
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Datum_der_Archivierung", OdbcType.VarChar, 255)
            param.Value = DatA
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Datum_der_Bearbeitung", OdbcType.VarChar, 255)
            param.Value = DatB
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Publikationsjahr", OdbcType.Text)
            param.Value = Publikationsjahr
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Bearbeiter_im_ZA", OdbcType.VarChar, 255)
            param.Value = Bearb
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Bemerkungen", OdbcType.VarChar, 255)
            param.Value = Bem
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Zugangsklasse", OdbcType.VarChar, 255)
            param.Value = Zugang
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Anzahl_Zeitreihen", OdbcType.Int, 11)
            param.Value = azr
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Zeitraum", OdbcType.VarChar, 32)
            param.Value = zeitraum
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@exportable", OdbcType.TinyInt, 4)
            param.Value = Exportierbar
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Fundort", OdbcType.Text)
            param.Value = Fundort
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Anmerkungsteil", OdbcType.Text)
            param.Value = Anmerkungsteil
            cmd.Parameters.Add(param)

            If Not DateiName Is Nothing Then
                param = New OdbcParameter("@datei_inhalt", OdbcType.Binary)
                DateiInhalt = GetDatei(DateiPfad)
                param.Value = DateiInhalt
                Debug.WriteLine(DateiInhalt)
                cmd.Parameters.Add(param)

                param = New OdbcParameter("@datei_name", OdbcType.VarChar, 255)
                param.IsNullable = True
                param.Value = DateiName
                cmd.Parameters.Add(param)

            End If

            msg = ""
            For i = 0 To cmd.Parameters.Count - 1
                msg += ControlChars.Cr + cmd.Parameters(i).ToString() + "" + cmd.Parameters(i).GetType.ToString + " " + cmd.Parameters(i).OdbcType.ToString + " " + _
                cmd.Parameters(i).Value.GetType.ToString()
            Next i

            Debug.WriteLine(msg)
            Debug.WriteLine(cmd.ToString)
            cmd.ExecuteNonQuery()
            Insert_Aka_Projekte = ID_Project

        Catch ex As Exception
            Insert_Aka_Projekte = Nothing
            Throw ex
        End Try
   End Function


   Private Sub Insert_Aka_Schluesselmaske()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Aka_Schluesselmaske"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Aka_Schluesselmaske (ID_HS, Name, ID_Projekt) Values (?, ?, ?);", conn)

         param = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Name", OdbcType.VarChar, 255, "Name")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@ID_Projekt", OdbcType.VarChar, 32, "ID_Projekt")
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd

         adapt.Update(ds, "DS_Aka_Schluesselmaske")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Private Sub Insert_Aka_Codes()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Aka_Codes"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Aka_Codes (ID_CodeKuerz, CodeBeschreibung, Zeichen) Values (?, ?, ?);", conn)

         param = New OdbcParameter("@ID_CodeKuerz", OdbcType.VarChar, 32, "ID_CodeKuerz")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@CodeBeschreibung", OdbcType.VarChar, 255, "CodeBeschreibung")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Zeichen", OdbcType.Int, 11, "Zeichen")
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd

         adapt.Update(ds, "DS_Aka_Codes")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub

   Private Sub Insert_Aka_SchluesselCode()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Aka_SchluesselCode"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Aka_SchluesselCode (ID_HS, ID_CodeKuerz, Position) Values (?, ?, ?);", conn)

         param = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@ID_CodeKuerz", OdbcType.VarChar, 32, "ID_CodeKuerz")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Position", OdbcType.Int, 11, "Position")
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd

         adapt.Update(ds, "DS_Aka_SchluesselCode")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub


   Private Sub Insert_Aka_CodeInhalt()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Aka_CodeInhalt"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Aka_CodeInhalt (ID_CodeKuerz, Code, CodeBezeichnung) Values (?, ?, ?);", conn)

         param = New OdbcParameter("@ID_CodeKuerz", OdbcType.VarChar, 32, "ID_CodeKuerz")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Code", OdbcType.VarChar, 255, "Code")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@CodeBezeichnung", OdbcType.VarChar, 255, "CodeBezeichnung")
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd
         adapt.Update(ds, "DS_Aka_CodeInhalt")
      Catch ex As Exception
         Throw ex
      End Try
   End Sub



   Private Sub Insert_Lit_ZR()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Lit_ZR"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Lit_ZR (ID_HS, Schluessel, Quelle, Anmerkung, Tabelle) Values (?, ?, ?, ?, ?);", conn)

         param = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Schluessel", OdbcType.VarChar, 255, "Schluessel")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Quelle", OdbcType.Text, 2147483647, "Quelle")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Anmerkung", OdbcType.Text, 2147483647, "Anmerkung")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Tabelle", OdbcType.VarChar, 255, "Tabelle")
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd

         outTabelle("DS_Lit_ZR")

         adapt.Update(ds, "DS_Lit_ZR")

      Catch ex As Exception
         Throw ex

      End Try
   End Sub


   Private Sub Insert_Daten__Aka()
      Dim param As OdbcParameter
      Dim cmd As OdbcCommand

      Try
         da.SPanel2.Text = "Insert_Daten__Aka"
         da.Refresh()

         cmd = New OdbcCommand("Insert Into Daten__Aka (ID_HS, Schluessel, Jahr_Sem, Data, Anmerkung) Values (?, ?, ?, ?, ?);", conn)

         param = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Schluessel", OdbcType.VarChar, 32, "Schluessel")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Jahr_Sem", OdbcType.VarChar, 11, "Jahr_Sem")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Data", OdbcType.Double, 11, "Data")
         cmd.Parameters.Add(param)

         param = New OdbcParameter("@Anmerkung", OdbcType.Text)
         param.SourceColumn = "Anmerkung"
         cmd.Parameters.Add(param)

         adapt.InsertCommand = cmd

         adapt.Update(ds, "DS_Daten__Aka")
      Catch ex As Exception
         Throw ex
      End Try

   End Sub


   Private Sub outTabelle(ByVal strTbl As String)
        Dim dt As System.Data.DataTable
      Dim dr As DataRow
      Dim i As Int16
      Dim j As Int16

      dt = ds.Tables(strTbl)
      For i = 0 To dt.Rows.Count - 1
         dr = dt.Rows(i)
         For j = 0 To dt.Columns.Count - 2
            If IsDBNull(dr.Item(j)) Then
               Debug.Write("Null".PadRight(39) & " ")
            Else
               Debug.Write(CStr(dr.Item(j)).PadRight(39) & " ")
            End If
         Next
         Debug.WriteLine("")
      Next

   End Sub


End Module
