Public Class Metadaten3
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
   Friend WithEvents cmdAbbruch As System.Windows.Forms.Button
   Friend WithEvents cmdBack As System.Windows.Forms.Button
   Friend WithEvents cmdWeiter As System.Windows.Forms.Button
   Friend WithEvents tbAnmerkungen As System.Windows.Forms.TextBox
   Friend WithEvents lblAnmerk As System.Windows.Forms.Label
   Friend WithEvents lblFundort As System.Windows.Forms.Label
   Friend WithEvents tbFundort As System.Windows.Forms.TextBox
    Friend WithEvents tbDatei As System.Windows.Forms.TextBox
    Friend WithEvents cmdDatei As System.Windows.Forms.Button
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents tbUntergliederung As System.Windows.Forms.TextBox
    Friend WithEvents lblUntergliederung As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdWeiter = New System.Windows.Forms.Button
        Me.tbAnmerkungen = New System.Windows.Forms.TextBox
        Me.lblAnmerk = New System.Windows.Forms.Label
        Me.lblFundort = New System.Windows.Forms.Label
        Me.cmdBack = New System.Windows.Forms.Button
        Me.tbFundort = New System.Windows.Forms.TextBox
        Me.cmdDatei = New System.Windows.Forms.Button
        Me.tbDatei = New System.Windows.Forms.TextBox
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.tbUntergliederung = New System.Windows.Forms.TextBox
        Me.lblUntergliederung = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdAbbruch
        '
        Me.cmdAbbruch.Location = New System.Drawing.Point(132, 380)
        Me.cmdAbbruch.Name = "cmdAbbruch"
        Me.cmdAbbruch.Size = New System.Drawing.Size(104, 23)
        Me.cmdAbbruch.TabIndex = 9
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdWeiter
        '
        Me.cmdWeiter.Location = New System.Drawing.Point(488, 380)
        Me.cmdWeiter.Name = "cmdWeiter"
        Me.cmdWeiter.Size = New System.Drawing.Size(104, 23)
        Me.cmdWeiter.TabIndex = 10
        Me.cmdWeiter.Text = "&Weiter >>"
        '
        'tbAnmerkungen
        '
        Me.tbAnmerkungen.Location = New System.Drawing.Point(132, 125)
        Me.tbAnmerkungen.MaxLength = 2147483647
        Me.tbAnmerkungen.Multiline = True
        Me.tbAnmerkungen.Name = "tbAnmerkungen"
        Me.tbAnmerkungen.Size = New System.Drawing.Size(460, 104)
        Me.tbAnmerkungen.TabIndex = 3
        Me.tbAnmerkungen.Text = ""
        '
        'lblAnmerk
        '
        Me.lblAnmerk.Location = New System.Drawing.Point(16, 125)
        Me.lblAnmerk.Name = "lblAnmerk"
        Me.lblAnmerk.Size = New System.Drawing.Size(96, 23)
        Me.lblAnmerk.TabIndex = 2
        Me.lblAnmerk.Text = "Anmerkungen"
        '
        'lblFundort
        '
        Me.lblFundort.Location = New System.Drawing.Point(16, 16)
        Me.lblFundort.Name = "lblFundort"
        Me.lblFundort.Size = New System.Drawing.Size(96, 23)
        Me.lblFundort.TabIndex = 0
        Me.lblFundort.Text = "Fundort"
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(8, 380)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(104, 23)
        Me.cmdBack.TabIndex = 8
        Me.cmdBack.Text = "<< &Zurück"
        '
        'tbFundort
        '
        Me.tbFundort.Location = New System.Drawing.Point(132, 16)
        Me.tbFundort.MaxLength = 2147483647
        Me.tbFundort.Multiline = True
        Me.tbFundort.Name = "tbFundort"
        Me.tbFundort.Size = New System.Drawing.Size(460, 104)
        Me.tbFundort.TabIndex = 1
        Me.tbFundort.Text = ""
        '
        'cmdDatei
        '
        Me.cmdDatei.Location = New System.Drawing.Point(8, 343)
        Me.cmdDatei.Name = "cmdDatei"
        Me.cmdDatei.Size = New System.Drawing.Size(104, 23)
        Me.cmdDatei.TabIndex = 6
        Me.cmdDatei.Text = "Datei auswählen.."
        '
        'tbDatei
        '
        Me.tbDatei.BackColor = System.Drawing.Color.White
        Me.tbDatei.Location = New System.Drawing.Point(132, 343)
        Me.tbDatei.Name = "tbDatei"
        Me.tbDatei.ReadOnly = True
        Me.tbDatei.Size = New System.Drawing.Size(460, 20)
        Me.tbDatei.TabIndex = 7
        Me.tbDatei.Text = ""
        '
        'OFD
        '
        Me.OFD.Filter = " pdf-Dateien (*.pdf)|*.pdf|Alle Dateien (*.*)|*.*"
        '
        'tbUntergliederung
        '
        Me.tbUntergliederung.Location = New System.Drawing.Point(132, 234)
        Me.tbUntergliederung.MaxLength = 2147483647
        Me.tbUntergliederung.Multiline = True
        Me.tbUntergliederung.Name = "tbUntergliederung"
        Me.tbUntergliederung.Size = New System.Drawing.Size(460, 104)
        Me.tbUntergliederung.TabIndex = 5
        Me.tbUntergliederung.Text = ""
        '
        'lblUntergliederung
        '
        Me.lblUntergliederung.Location = New System.Drawing.Point(16, 234)
        Me.lblUntergliederung.Name = "lblUntergliederung"
        Me.lblUntergliederung.Size = New System.Drawing.Size(96, 46)
        Me.lblUntergliederung.TabIndex = 4
        Me.lblUntergliederung.Text = "Sachliche Untergliederung der Datentabellen"
        '
        'Metadaten3
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 423)
        Me.Controls.Add(Me.tbUntergliederung)
        Me.Controls.Add(Me.lblUntergliederung)
        Me.Controls.Add(Me.tbDatei)
        Me.Controls.Add(Me.tbFundort)
        Me.Controls.Add(Me.tbAnmerkungen)
        Me.Controls.Add(Me.cmdDatei)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdWeiter)
        Me.Controls.Add(Me.lblAnmerk)
        Me.Controls.Add(Me.lblFundort)
        Me.Name = "Metadaten3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Try
            md2.Show()
            Me.Hide()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub cmdWeiter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWeiter.Click
        Try
            Fundort = tbFundort.Text
            Anmerkungsteil = tbAnmerkungen.Text
            Untergliederung = tbUntergliederung.Text
            '  DateiName = tbDatei.Text

            If status = Globals.tStatus.onMetadataUpdate Then
                speichern()
            Else
                If IsNothing(da) Then
                    da = New Dateiauswahl
                End If
                Me.Hide()
                da.Show()
            End If
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub cmdAbbruch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbbruch.Click
        closeForms()
    End Sub


    Private Sub Metadaten1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        If Not onAbbr Then
            onAbbr = True
            If Not IsNothing(anm) Then
                anm.Close()
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


    Private Sub Metadaten3_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If status = Globals.tStatus.onMetadataUpdate Then
            cmdWeiter.Text = "&Fertigstellen"

            Me.tbFundort.Text = Fundort
            Me.tbAnmerkungen.Text = Anmerkungsteil
            Me.tbUntergliederung.Text = Untergliederung
            Me.tbDatei.Text = DateiName
        End If
    End Sub

    Private Sub speichern()
        Try
            Me.Cursor = Cursors.WaitCursor
            If DBUpdateMetadata.Update_Aka_Projekte() Then
                Me.Cursor = Cursors.Default
                HistatLog.WriteLine(Now & " Die Projektbeschreibung wurde geändert." & vbCrLf)
                MsgBox("Die Projektbeschreibung wurde geändert!", MsgBoxStyle.Information)
                sf.Show()
            End If
            Me.Close()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub cmdDatei_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDatei.Click
        Dim dr As DialogResult
        Dim i As Int16
        Try
            With OFD
                dr = .ShowDialog()
                If dr = DialogResult.OK Then
                    tbDatei.Text = .FileName
                    DateiPfad = .FileName
                    i = tbDatei.Text.LastIndexOf("\")
                    DateiName = tbDatei.Text.Substring(i + 1)
                Else
                End If
            End With
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try

    End Sub


    Private Sub tbDatei_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDatei.Click
        If MsgBox("Wollen Sie Die Datei wirklich aus der Projektbeschreibung entfernen?", MsgBoxStyle.YesNo, "Histat") = vbYes Then
            DateiPfad = Nothing
            DateiName = Nothing
            tbDatei.Text = Nothing
        End If
    End Sub

    
End Class
