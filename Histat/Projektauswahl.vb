Public Class Projektauswahl
    Inherits System.Windows.Forms.Form

#Region " Vom Windows Form Designer generierter Code "

    Public Sub New()
        MyBase.New()

        ' Dieser Aufruf ist für den Windows Form-Designer erforderlich.
        InitializeComponent()

        ' Initialisierungen nach dem Aufruf InitializeComponent() hinzufügen

    End Sub

    ' Die Form überschreibt den Löschvorgang der Basisklasse, um Komponenten zu bereinigen.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Für Windows Form-Designer erforderlich
    Private components As System.ComponentModel.IContainer

    'HINWEIS: Die folgende Prozedur ist für den Windows Form-Designer erforderlich
    'Sie kann mit dem Windows Form-Designer modifiziert werden.
    'Verwenden Sie nicht den Code-Editor zur Bearbeitung.
    Friend WithEvents lblProjekt As System.Windows.Forms.Label
    Friend WithEvents lblAutor As System.Windows.Forms.Label
    Friend WithEvents lblBeschreibung As System.Windows.Forms.Label
    Friend WithEvents cmdAbbruch As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents tbAutor As System.Windows.Forms.TextBox
    Friend WithEvents tbBeschreibung As System.Windows.Forms.TextBox
    Friend WithEvents cbProjekt As System.Windows.Forms.ComboBox
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents tbchDate As System.Windows.Forms.TextBox
    Friend WithEvents tbAnzZr As System.Windows.Forms.TextBox
    Friend WithEvents cbID_Zeit As System.Windows.Forms.ComboBox
    Friend WithEvents lblID_Zeit As System.Windows.Forms.Label
    Friend WithEvents lblZr As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblProjekt = New System.Windows.Forms.Label
        Me.tbAutor = New System.Windows.Forms.TextBox
        Me.lblAutor = New System.Windows.Forms.Label
        Me.tbBeschreibung = New System.Windows.Forms.TextBox
        Me.lblBeschreibung = New System.Windows.Forms.Label
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cbProjekt = New System.Windows.Forms.ComboBox
        Me.lblDate = New System.Windows.Forms.Label
        Me.tbchDate = New System.Windows.Forms.TextBox
        Me.tbAnzZr = New System.Windows.Forms.TextBox
        Me.lblZr = New System.Windows.Forms.Label
        Me.cbID_Zeit = New System.Windows.Forms.ComboBox
        Me.lblID_Zeit = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lblProjekt
        '
        Me.lblProjekt.Location = New System.Drawing.Point(16, 16)
        Me.lblProjekt.Name = "lblProjekt"
        Me.lblProjekt.Size = New System.Drawing.Size(56, 23)
        Me.lblProjekt.TabIndex = 0
        Me.lblProjekt.Text = "Projekt"
        '
        'tbAutor
        '
        Me.tbAutor.Enabled = False
        Me.tbAutor.Location = New System.Drawing.Point(108, 48)
        Me.tbAutor.MaxLength = 255
        Me.tbAutor.Name = "tbAutor"
        Me.tbAutor.Size = New System.Drawing.Size(484, 20)
        Me.tbAutor.TabIndex = 3
        '
        'lblAutor
        '
        Me.lblAutor.Location = New System.Drawing.Point(16, 48)
        Me.lblAutor.Name = "lblAutor"
        Me.lblAutor.Size = New System.Drawing.Size(56, 23)
        Me.lblAutor.TabIndex = 2
        Me.lblAutor.Text = "Autor"
        '
        'tbBeschreibung
        '
        Me.tbBeschreibung.Enabled = False
        Me.tbBeschreibung.Location = New System.Drawing.Point(108, 80)
        Me.tbBeschreibung.MaxLength = 2147483647
        Me.tbBeschreibung.Multiline = True
        Me.tbBeschreibung.Name = "tbBeschreibung"
        Me.tbBeschreibung.Size = New System.Drawing.Size(484, 248)
        Me.tbBeschreibung.TabIndex = 5
        '
        'lblBeschreibung
        '
        Me.lblBeschreibung.Location = New System.Drawing.Point(16, 88)
        Me.lblBeschreibung.Name = "lblBeschreibung"
        Me.lblBeschreibung.Size = New System.Drawing.Size(80, 23)
        Me.lblBeschreibung.TabIndex = 4
        Me.lblBeschreibung.Text = "Beschreibung"
        '
        'cmdAbbruch
        '
        Me.cmdAbbruch.Location = New System.Drawing.Point(16, 403)
        Me.cmdAbbruch.Name = "cmdAbbruch"
        Me.cmdAbbruch.Size = New System.Drawing.Size(104, 23)
        Me.cmdAbbruch.TabIndex = 9
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(488, 403)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(104, 23)
        Me.cmdOK.TabIndex = 10
        Me.cmdOK.Text = "&Löschen"
        '
        'cbProjekt
        '
        Me.cbProjekt.Location = New System.Drawing.Point(108, 16)
        Me.cbProjekt.Name = "cbProjekt"
        Me.cbProjekt.Size = New System.Drawing.Size(484, 21)
        Me.cbProjekt.TabIndex = 1
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(16, 371)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(88, 23)
        Me.lblDate.TabIndex = 6
        Me.lblDate.Text = "Letzte Änderung"
        '
        'tbchDate
        '
        Me.tbchDate.Enabled = False
        Me.tbchDate.Location = New System.Drawing.Point(108, 371)
        Me.tbchDate.Name = "tbchDate"
        Me.tbchDate.Size = New System.Drawing.Size(212, 20)
        Me.tbchDate.TabIndex = 7
        '
        'tbAnzZr
        '
        Me.tbAnzZr.Enabled = False
        Me.tbAnzZr.Location = New System.Drawing.Point(512, 371)
        Me.tbAnzZr.Name = "tbAnzZr"
        Me.tbAnzZr.Size = New System.Drawing.Size(80, 20)
        Me.tbAnzZr.TabIndex = 8
        '
        'lblZr
        '
        Me.lblZr.Location = New System.Drawing.Point(448, 371)
        Me.lblZr.Name = "lblZr"
        Me.lblZr.Size = New System.Drawing.Size(56, 23)
        Me.lblZr.TabIndex = 10
        Me.lblZr.Text = "Zeitreihen"
        '
        'cbID_Zeit
        '
        Me.cbID_Zeit.FormattingEnabled = True
        Me.cbID_Zeit.Location = New System.Drawing.Point(108, 340)
        Me.cbID_Zeit.Name = "cbID_Zeit"
        Me.cbID_Zeit.Size = New System.Drawing.Size(212, 21)
        Me.cbID_Zeit.TabIndex = 11
        '
        'lblID_Zeit
        '
        Me.lblID_Zeit.AutoSize = True
        Me.lblID_Zeit.Location = New System.Drawing.Point(16, 340)
        Me.lblID_Zeit.Name = "lblID_Zeit"
        Me.lblID_Zeit.Size = New System.Drawing.Size(25, 13)
        Me.lblID_Zeit.TabIndex = 12
        Me.lblID_Zeit.Text = "Zeit"
        '
        'Projektauswahl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 438)
        Me.Controls.Add(Me.lblID_Zeit)
        Me.Controls.Add(Me.cbID_Zeit)
        Me.Controls.Add(Me.tbAnzZr)
        Me.Controls.Add(Me.tbchDate)
        Me.Controls.Add(Me.tbBeschreibung)
        Me.Controls.Add(Me.tbAutor)
        Me.Controls.Add(Me.lblZr)
        Me.Controls.Add(Me.lblDate)
        Me.Controls.Add(Me.cbProjekt)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblBeschreibung)
        Me.Controls.Add(Me.lblAutor)
        Me.Controls.Add(Me.lblProjekt)
        Me.MaximizeBox = False
        Me.Name = "Projektauswahl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region


    Private conn As System.Data.Odbc.OdbcConnection
    Private adapt As OdbcDataAdapter
    Private ds As DataSet


    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try
            If ID_Projekt = "-1" Then
                MsgBox("Bitte wählen Sie ein Projekt aus der Auswahlbox aus!", MsgBoxStyle.Exclamation, "Histat")
            ElseIf status = Globals.tStatus.onDelete Then
                loeschen()
            ElseIf status = Globals.tStatus.onMetadataUpdate Then
                If Me.cbProjekt.Text = "" Then
                    MsgBox("Bitte geben Sie einen Projektnamen an!", MsgBoxStyle.Exclamation, "Histat")
                    Exit Sub
                End If
                Projektname = Me.cbProjekt.Text
                Projektautor = Me.tbAutor.Text
                Projektbeschreibung = Me.tbBeschreibung.Text
                Me.Hide()
                md1 = New Metadaten1
                md1.Show()
            Else               ' status = Globals.tStatus.onDataImport 
                ID_Projekt = cbProjekt.SelectedValue
                Me.Hide()
                da = New Dateiauswahl
                da.Show()
            End If
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub Projektauswahl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sSql As String
        Dim dt As System.Data.DataTable
        Dim key(0) As DataColumn
        Try
            If status = Globals.tStatus.onDataImport Then
                cmdOK.Text = "&Weiter"
            End If
            If status = Globals.tStatus.onMetadataUpdate Then
                cmdOK.Text = "&Weiter"
                Me.tbAutor.Enabled = True
                Me.tbBeschreibung.Enabled = True
                sSql = "SELECT ID_Projekt, ID_Zeit, Projektautor, Projektname, Projektbeschreibung, Veroeffentlichung, Publikationsjahr, " & _
                        "Untersuchungsgebiet, Quellen, Untergliederung, ZA_Studiennummer, Datum_der_Archivierung, Datum_der_Bearbeitung, " & _
                        "Bearbeiter_im_ZA, Bemerkungen, Zugangsklasse, Fundort, Anmerkungsteil, Anzahl_Zeitreihen, " & _
                        "Zeitraum, exportable, datei_name, chdate FROM Aka_Projekte " & _
                        "WHERE ID_Thema = " & ID_Thema & ";"
            Else
                sSql = "SELECT ID_Projekt, ID_Zeit, Projektautor, Projektname, Projektbeschreibung, Anzahl_Zeitreihen, exportable, chdate FROM Aka_Projekte " & _
                    "WHERE ID_Thema = " & ID_Thema & ";"
            End If

            conn = New System.Data.Odbc.OdbcConnection(conStr)
            ds = New DataSet
            conn.Open()

            Debug.WriteLine(sSql)
            adapt = New OdbcDataAdapter(sSql, conn)
            adapt.Fill(ds, "DS_Aka_Projekte")
            dt = ds.Tables("DS_Aka_Projekte")
            key(0) = dt.Columns(0)
            dt.PrimaryKey = key
            '          conn.Close()

            With cbProjekt
                .DataSource = ds.Tables("DS_Aka_Projekte")
                .DisplayMember = "Projektname"
                .ValueMember = "ID_Projekt"
                .SelectedIndex = -1
            End With

            tbAutor.Text = ""
            tbBeschreibung.Text = ""

            adapt = Nothing
            sSql = "SELECT ID_Zeit, Zeit FROM Aka_Zeiten ORDER BY Position;"
            adapt = New OdbcDataAdapter(sSql, conn)
            adapt.Fill(ds, "DS_Aka_Zeiten")
            dt = ds.Tables("DS_Aka_Zeiten")
            key(0) = dt.Columns(0)
            dt.PrimaryKey = key

            With cbID_Zeit
                .DataSource = ds.Tables("DS_Aka_Zeiten")
                .DisplayMember = "Zeit"
                .ValueMember = "ID_Zeit"
                .SelectedIndex = -1
            End With

            cbID_Zeit.Text = ""
            tbchDate.Text = ""
            tbAnzZr.Text = ""

            conn.Close()
            adapt = Nothing
            conn = Nothing

        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub



    Private Sub cbProjekt_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbProjekt.DropDown
        Try
            With cbProjekt
                .SelectedIndex = .FindString(.Text)
            End With

        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub cbProjekt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbProjekt.SelectedIndexChanged
        Dim dr As DataRow
        Try
            With cbProjekt
                If Not IsNothing(.SelectedItem) Then
                    ID_Projekt = .SelectedValue.ToString
                    dr = .SelectedItem.row
                    tbAutor.Text = dr.Item("Projektautor")
                    tbBeschreibung.Text = dr.Item("Projektbeschreibung")
                    tbchDate.Text = dr.Item("chdate")
                    tbAnzZr.Text = dr.Item("Anzahl_Zeitreihen")
                    Exportierbar = dr.Item("exportable")
                    If status = Globals.tStatus.onMetadataUpdate Then
                        getMetadata(dr)
                    End If
                    cbID_Zeit.SelectedValue = dr.Item("ID_Zeit")
                Else
                    ID_Projekt = "-1"
                End If
            End With



        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")

        End Try
    End Sub

    Private Sub Projektauswahl_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        ds = Nothing
        adapt = Nothing
        If Not IsNothing(conn) Then
            If Not conn.State = ConnectionState.Closed Then
                conn.Close()
            End If
            conn = Nothing
        End If
    End Sub

    Private Sub cmdAbbruch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAbbruch.Click
        Me.Close()
    End Sub

    Private Sub Projektauswahl_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        sf.Show()
    End Sub

    Private Sub loeschen()
        Dim strProjekt As String
        Try
            strProjekt = cbProjekt.Text
            If MsgBox("Möchten Sie das Projekt '" & strProjekt & "' wirklich löschen?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                del_Projekt(cbProjekt.SelectedValue)
                Me.Cursor = Cursors.Default
                HistatLog.WriteLine(Now & "Das Projekt " & strProjekt & " wurde erfolgreich gelöscht." & vbCrLf)
                MsgBox("Das Projekt " & strProjekt & " wurde erfolgreich gelöscht.", MsgBoxStyle.Information)
                sf.Show()
                Me.Close()
            End If
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub getMetadata(ByVal dr As DataRow)
        ID_Zeit = dr.Item("ID_Zeit")
        Projektautor = dr.Item("Projektautor")
        Projektname = dr.Item("Projektname")
        Projektbeschreibung = dr.Item("Projektbeschreibung")
        Veroeff = dr.Item("Veroeffentlichung")
        Untersuch = dr.Item("Untersuchungsgebiet")
        Quellen = dr.Item("Quellen")
        ZAStd = dr.Item("ZA_Studiennummer")
        DatA = dr.Item("Datum_der_Archivierung")
        DatB = dr.Item("Datum_der_Bearbeitung")
        Publikationsjahr = dr.Item("Publikationsjahr")
        Bearb = dr.Item("Bearbeiter_im_ZA")
        Bem = dr.Item("Bemerkungen")
        Zugang = dr.Item("Zugangsklasse")
        Fundort = dr.Item("Fundort")
        Anmerkungsteil = dr.Item("Anmerkungsteil")
        Exportierbar = dr.Item("exportable")
        DateiName = IIf(IsDBNull(dr.Item("datei_name")), Nothing, dr.Item("datei_name"))
        Untergliederung = dr.Item("Untergliederung")
    End Sub



End Class

