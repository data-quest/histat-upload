Public Class Metadaten1
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
   Friend WithEvents lblVeroeff As System.Windows.Forms.Label
   Friend WithEvents tbUntersuch As System.Windows.Forms.TextBox
   Friend WithEvents lblUntersuch As System.Windows.Forms.Label
   Friend WithEvents lblQuellen As System.Windows.Forms.Label
   Friend WithEvents tbVeroeff As System.Windows.Forms.TextBox
   Friend WithEvents cmdBack As System.Windows.Forms.Button
    Friend WithEvents cmdWeiter As System.Windows.Forms.Button
    Friend WithEvents tbPublikationsjahr As System.Windows.Forms.TextBox
    Friend WithEvents lblPublikationsjahr As System.Windows.Forms.Label
    Friend WithEvents tbQuellen As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdWeiter = New System.Windows.Forms.Button
        Me.tbUntersuch = New System.Windows.Forms.TextBox
        Me.lblUntersuch = New System.Windows.Forms.Label
        Me.lblQuellen = New System.Windows.Forms.Label
        Me.lblVeroeff = New System.Windows.Forms.Label
        Me.cmdBack = New System.Windows.Forms.Button
        Me.tbVeroeff = New System.Windows.Forms.TextBox
        Me.tbQuellen = New System.Windows.Forms.TextBox
        Me.tbPublikationsjahr = New System.Windows.Forms.TextBox
        Me.lblPublikationsjahr = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'cmdAbbruch
        '
        Me.cmdAbbruch.Location = New System.Drawing.Point(132, 420)
        Me.cmdAbbruch.Name = "cmdAbbruch"
        Me.cmdAbbruch.Size = New System.Drawing.Size(104, 23)
        Me.cmdAbbruch.TabIndex = 7
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdWeiter
        '
        Me.cmdWeiter.Location = New System.Drawing.Point(488, 420)
        Me.cmdWeiter.Name = "cmdWeiter"
        Me.cmdWeiter.Size = New System.Drawing.Size(104, 23)
        Me.cmdWeiter.TabIndex = 8
        Me.cmdWeiter.Text = "&Weiter >>"
        '
        'tbUntersuch
        '
        Me.tbUntersuch.Location = New System.Drawing.Point(132, 171)
        Me.tbUntersuch.MaxLength = 2147483647
        Me.tbUntersuch.Multiline = True
        Me.tbUntersuch.Name = "tbUntersuch"
        Me.tbUntersuch.Size = New System.Drawing.Size(460, 104)
        Me.tbUntersuch.TabIndex = 2
        '
        'lblUntersuch
        '
        Me.lblUntersuch.Location = New System.Drawing.Point(15, 175)
        Me.lblUntersuch.Name = "lblUntersuch"
        Me.lblUntersuch.Size = New System.Drawing.Size(113, 23)
        Me.lblUntersuch.TabIndex = 2
        Me.lblUntersuch.Text = "Untersuchungsgebiet"
        '
        'lblQuellen
        '
        Me.lblQuellen.Location = New System.Drawing.Point(15, 292)
        Me.lblQuellen.Name = "lblQuellen"
        Me.lblQuellen.Size = New System.Drawing.Size(56, 23)
        Me.lblQuellen.TabIndex = 4
        Me.lblQuellen.Text = "Quellen"
        '
        'lblVeroeff
        '
        Me.lblVeroeff.Location = New System.Drawing.Point(15, 18)
        Me.lblVeroeff.Name = "lblVeroeff"
        Me.lblVeroeff.Size = New System.Drawing.Size(96, 23)
        Me.lblVeroeff.TabIndex = 0
        Me.lblVeroeff.Text = "Veröffentlichung "
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(8, 420)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(104, 23)
        Me.cmdBack.TabIndex = 6
        Me.cmdBack.Text = "<< &Zurück"
        '
        'tbVeroeff
        '
        Me.tbVeroeff.Location = New System.Drawing.Point(132, 14)
        Me.tbVeroeff.MaxLength = 2147483647
        Me.tbVeroeff.Multiline = True
        Me.tbVeroeff.Name = "tbVeroeff"
        Me.tbVeroeff.Size = New System.Drawing.Size(460, 104)
        Me.tbVeroeff.TabIndex = 1
        '
        'tbQuellen
        '
        Me.tbQuellen.Location = New System.Drawing.Point(132, 288)
        Me.tbQuellen.MaxLength = 2147483647
        Me.tbQuellen.Multiline = True
        Me.tbQuellen.Name = "tbQuellen"
        Me.tbQuellen.Size = New System.Drawing.Size(460, 104)
        Me.tbQuellen.TabIndex = 3
        '
        'tbPublikationsjahr
        '
        Me.tbPublikationsjahr.Location = New System.Drawing.Point(132, 125)
        Me.tbPublikationsjahr.Name = "tbPublikationsjahr"
        Me.tbPublikationsjahr.Size = New System.Drawing.Size(460, 20)
        Me.tbPublikationsjahr.TabIndex = 9
        '
        'lblPublikationsjahr
        '
        Me.lblPublikationsjahr.AutoSize = True
        Me.lblPublikationsjahr.Location = New System.Drawing.Point(15, 131)
        Me.lblPublikationsjahr.Name = "lblPublikationsjahr"
        Me.lblPublikationsjahr.Size = New System.Drawing.Size(81, 13)
        Me.lblPublikationsjahr.TabIndex = 10
        Me.lblPublikationsjahr.Text = "Publikationsjahr"
        '
        'Metadaten1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(613, 449)
        Me.Controls.Add(Me.lblPublikationsjahr)
        Me.Controls.Add(Me.tbPublikationsjahr)
        Me.Controls.Add(Me.tbQuellen)
        Me.Controls.Add(Me.tbVeroeff)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdWeiter)
        Me.Controls.Add(Me.tbUntersuch)
        Me.Controls.Add(Me.lblUntersuch)
        Me.Controls.Add(Me.lblQuellen)
        Me.Controls.Add(Me.lblVeroeff)
        Me.Name = "Metadaten1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region




    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Try
            If status = Globals.tStatus.onMetadataUpdate Then
                pa.Show()
            Else
                anm.Show()
            End If
            Me.Hide()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub cmdWeiter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWeiter.Click
        Try
            Veroeff = tbVeroeff.Text
            Publikationsjahr = tbPublikationsjahr.Text
            Untersuch = tbUntersuch.Text
            Quellen = tbQuellen.Text

            Me.Hide()
            If IsNothing(md2) Then
                md2 = New Metadaten2()
            End If
            md2.Show()

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


   Private Sub Metadaten1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
      If status = Globals.tStatus.onMetadataUpdate Then
            Me.tbVeroeff.Text = Veroeff
            Me.tbPublikationsjahr.Text = Publikationsjahr
            Me.tbUntersuch.Text = Untersuch
         Me.tbQuellen.Text = Quellen
        End If

   End Sub

    Private Sub tbPublikationsjahr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPublikationsjahr.TextChanged

    End Sub
End Class
