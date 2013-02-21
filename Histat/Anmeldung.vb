Public Class Anmeldung
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
   Friend WithEvents tbProjekt As System.Windows.Forms.TextBox
   Friend WithEvents lblAutor As System.Windows.Forms.Label
   Friend WithEvents lblBeschreibung As System.Windows.Forms.Label
   Friend WithEvents cmdAbbruch As System.Windows.Forms.Button
   Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents tbAutor As System.Windows.Forms.TextBox
    Friend WithEvents cbID_Zeit As System.Windows.Forms.ComboBox
    Friend WithEvents lblID_Zeit As System.Windows.Forms.Label
    Friend WithEvents tbBeschreibung As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblProjekt = New System.Windows.Forms.Label
        Me.tbProjekt = New System.Windows.Forms.TextBox
        Me.tbAutor = New System.Windows.Forms.TextBox
        Me.lblAutor = New System.Windows.Forms.Label
        Me.tbBeschreibung = New System.Windows.Forms.TextBox
        Me.lblBeschreibung = New System.Windows.Forms.Label
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
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
        'tbProjekt
        '
        Me.tbProjekt.Location = New System.Drawing.Point(104, 16)
        Me.tbProjekt.MaxLength = 255
        Me.tbProjekt.Name = "tbProjekt"
        Me.tbProjekt.Size = New System.Drawing.Size(488, 20)
        Me.tbProjekt.TabIndex = 1
        '
        'tbAutor
        '
        Me.tbAutor.Location = New System.Drawing.Point(104, 48)
        Me.tbAutor.MaxLength = 255
        Me.tbAutor.Name = "tbAutor"
        Me.tbAutor.Size = New System.Drawing.Size(488, 20)
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
        Me.tbBeschreibung.Location = New System.Drawing.Point(104, 88)
        Me.tbBeschreibung.MaxLength = 2147483647
        Me.tbBeschreibung.Multiline = True
        Me.tbBeschreibung.Name = "tbBeschreibung"
        Me.tbBeschreibung.Size = New System.Drawing.Size(488, 256)
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
        Me.cmdAbbruch.TabIndex = 6
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdOK
        '
        Me.cmdOK.Location = New System.Drawing.Point(488, 403)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(104, 23)
        Me.cmdOK.TabIndex = 7
        Me.cmdOK.Text = "&Weiter >>"
        '
        'cbID_Zeit
        '
        Me.cbID_Zeit.FormattingEnabled = True
        Me.cbID_Zeit.Location = New System.Drawing.Point(104, 362)
        Me.cbID_Zeit.Name = "cbID_Zeit"
        Me.cbID_Zeit.Size = New System.Drawing.Size(212, 21)
        Me.cbID_Zeit.TabIndex = 8
        '
        'lblID_Zeit
        '
        Me.lblID_Zeit.AutoSize = True
        Me.lblID_Zeit.Location = New System.Drawing.Point(16, 365)
        Me.lblID_Zeit.Name = "lblID_Zeit"
        Me.lblID_Zeit.Size = New System.Drawing.Size(25, 13)
        Me.lblID_Zeit.TabIndex = 9
        Me.lblID_Zeit.Text = "Zeit"
        '
        'Anmeldung
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 438)
        Me.Controls.Add(Me.lblID_Zeit)
        Me.Controls.Add(Me.cbID_Zeit)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.tbBeschreibung)
        Me.Controls.Add(Me.tbAutor)
        Me.Controls.Add(Me.tbProjekt)
        Me.Controls.Add(Me.lblBeschreibung)
        Me.Controls.Add(Me.lblAutor)
        Me.Controls.Add(Me.lblProjekt)
        Me.Name = "Anmeldung"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Try

            Projektname = tbProjekt.Text
            If Projektname = "" Then
                MsgBox("Bitte geben Sie einen Projektnamen an!", , "Histat-Import")
                tbProjekt.Focus()
                Exit Sub
            End If

            Projektautor = tbAutor.Text
            If Projektautor = "" Then
                MsgBox("Bitte geben Sie einen Projektautor an!", , "Histat-Import")
                tbAutor.Focus()
                Exit Sub
            End If

            Projektbeschreibung = tbBeschreibung.Text
            If Projektbeschreibung = "" Then
                MsgBox("Bitte geben Sie eine Projektbeschreibung an!", , "Histat-Import")
                tbBeschreibung.Focus()
                Exit Sub
            End If

            If cbID_Zeit.SelectedValue Is Nothing Then
                MsgBox("Bitte geben Sie eine Zeit an!", , "Histat-Import")
                cbID_Zeit.Focus()
                Exit Sub
            End If
            ID_Zeit = cbID_Zeit.SelectedValue

            If IsNothing(md1) Then
                md1 = New Metadaten1()
            End If
            Me.Hide()
            md1.Show()

        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub cmdAbbruch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbbruch.Click
        closeForms()
    End Sub


    Private Sub Anmeldung_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        If Not onAbbr Then
            onAbbr = True
            If Not IsNothing(md1) Then
                md1.Close()
            End If
            If Not IsNothing(md2) Then
                md2.Close()
            End If
            If Not IsNothing(da) Then
                da.Close()
            End If
            sf.Show()
        End If

    End Sub


    Private Sub Anmeldung_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim conn As System.Data.Odbc.OdbcConnection
        Dim adapt As OdbcDataAdapter
        Dim ds As DataSet

        Dim sSql As String
        Dim dt As System.Data.DataTable
        Dim key(0) As DataColumn
        Try
            conn = New System.Data.Odbc.OdbcConnection(conStr)
            ds = New DataSet
            conn.Open()

            sSql = "SELECT ID_Zeit, Zeit FROM Aka_Zeiten ORDER BY Position;"
            adapt = New OdbcDataAdapter(sSql, conn)
            adapt.Fill(ds, "DS_Aka_Zeiten")
            dt = ds.Tables("DS_Aka_Zeiten")
            key(0) = dt.Columns(0)
            dt.PrimaryKey = key
            conn.Close()

            With cbID_Zeit
                .DisplayMember = "Zeit"
                .ValueMember = "ID_Zeit"
                .DataSource = ds.Tables("DS_Aka_Zeiten")
                .SelectedIndex = -1
            End With

            adapt = Nothing
            ds = Nothing
            conn = Nothing

        Catch ex As Exception
            Throw ex
        End Try

    End Sub


 End Class

