


Module DBUpdateMetadata

    Private conn As System.Data.Odbc.OdbcConnection


    Public Function Update_Aka_Projekte() As Boolean
        '   liefert bei erfolgreichem Eintrag des Projekts  true zurück, false sonst.
        Dim cmd As OdbcCommand
        Dim param As OdbcParameter
        Dim DateiInhalt() As Byte
        Dim i As Int16
        Dim msg As String

        Try
            conn = New System.Data.Odbc.OdbcConnection(conStr)
            conn.Open()
            If DateiName Is Nothing Then
                cmd = New OdbcCommand("UPDATE Aka_Projekte Set ID_Zeit = ?, Projektautor = ?, Projektname = ?, Projektbeschreibung = ?, " & _
                    "Veroeffentlichung = ?, Untersuchungsgebiet = ?, Quellen = ?, Untergliederung = ?, ZA_Studiennummer = ? ," & _
                    "Datum_der_Archivierung = ?, Datum_der_Bearbeitung = ?, Publikationsjahr = ?, Bearbeiter_im_ZA = ?, Bemerkungen = ?, " & _
                    "Zugangsklasse = ?, exportable = ?, Fundort = ?, Anmerkungsteil = ?, datei_inhalt = Null, datei_name = Null " & _
                    "WHERE ID_Projekt = ?;", conn)
            ElseIf DateiPfad Is Nothing Then
                cmd = New OdbcCommand("UPDATE Aka_Projekte Set ID_Zeit = ?, Projektautor = ?, Projektname = ?, Projektbeschreibung = ?, " & _
                    "Veroeffentlichung = ?, Untersuchungsgebiet = ?, Quellen = ?, Untergliederung = ?, ZA_Studiennummer = ? ," & _
                    "Datum_der_Archivierung = ?, Datum_der_Bearbeitung = ?, Publikationsjahr = ?, Bearbeiter_im_ZA = ?, Bemerkungen = ?, " & _
                    "Zugangsklasse = ?, exportable = ?, Fundort = ?, Anmerkungsteil = ? " & _
                    "WHERE ID_Projekt = ?;", conn)
            Else

                cmd = New OdbcCommand("UPDATE Aka_Projekte Set ID_Zeit = ?, Projektautor = ?, Projektname = ?, Projektbeschreibung = ?, " & _
                    "Veroeffentlichung = ?, Untersuchungsgebiet = ?, Quellen = ?, Untergliederung = ?, ZA_Studiennummer = ? ," & _
                    "Datum_der_Archivierung = ?, Datum_der_Bearbeitung = ?, Publikationsjahr = ?, Bearbeiter_im_ZA = ?, Bemerkungen = ?, " & _
                    "Zugangsklasse = ?, exportable = ?, Fundort = ?, Anmerkungsteil = ?, datei_inhalt = ?, datei_name = ? " & _
                    "WHERE ID_Projekt = ?;", conn)

            End If

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

            param = New OdbcParameter("@Publikationsjahr", OdbcType.Text)       '???????????
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

            param = New OdbcParameter("@exportable", OdbcType.TinyInt, 4)
            param.Value = Exportierbar
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Fundort", OdbcType.Text)
            param.Value = Fundort
            cmd.Parameters.Add(param)

            param = New OdbcParameter("@Anmerkungsteil", OdbcType.Text)
            param.Value = Anmerkungsteil
            cmd.Parameters.Add(param)

            If Not (DateiName Is Nothing Or DateiPfad Is Nothing) Then

                DateiInhalt = GetDatei(DateiPfad)
                param = New OdbcParameter("@datei_inhalt", OdbcType.Binary)
                param.Value = DateiInhalt
                Debug.WriteLine(DateiInhalt)
                cmd.Parameters.Add(param)

                param = New OdbcParameter("@datei_name", OdbcType.VarChar, 255)
                param.Value = DateiName
                cmd.Parameters.Add(param)
            End If

            param = New OdbcParameter("@ID_Projekt", OdbcType.VarChar, 32)
            param.Value = ID_Projekt
            cmd.Parameters.Add(param)

            msg = ""
            For i = 0 To cmd.Parameters.Count - 1
                msg += ControlChars.Cr + cmd.Parameters(i).ToString() + "" + cmd.Parameters(i).GetType.ToString + " " + cmd.Parameters(i).OdbcType.ToString + " " + _
                cmd.Parameters(i).Value.GetType.ToString()
            Next i

            Debug.WriteLine(msg)
            Debug.WriteLine(cmd.ToString)
            cmd.ExecuteNonQuery()
            conn.Close()
            Update_Aka_Projekte = True

        Catch ex As Exception
            Update_Aka_Projekte = False
            Throw ex
        End Try
    End Function
End Module
