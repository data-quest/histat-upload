Public Class Metadaten2
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
   Friend WithEvents tbZugang As System.Windows.Forms.TextBox
   Friend WithEvents lblZugang As System.Windows.Forms.Label
   Friend WithEvents tbBem As System.Windows.Forms.TextBox
   Friend WithEvents lblBem As System.Windows.Forms.Label
   Friend WithEvents lblBearb As System.Windows.Forms.Label
   Friend WithEvents tbDatA As System.Windows.Forms.TextBox
   Friend WithEvents tbDatB As System.Windows.Forms.TextBox
   Friend WithEvents cmdBack As System.Windows.Forms.Button
   Friend WithEvents lblDatA As System.Windows.Forms.Label
   Friend WithEvents lblDatB As System.Windows.Forms.Label
   Friend WithEvents cmdWeiter As System.Windows.Forms.Button
   Friend WithEvents tbBearb As System.Windows.Forms.TextBox
   Friend WithEvents tbZAStd As System.Windows.Forms.TextBox
   Friend WithEvents lblZA As System.Windows.Forms.Label
   Friend WithEvents chkExport As System.Windows.Forms.CheckBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdBack = New System.Windows.Forms.Button
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdWeiter = New System.Windows.Forms.Button
        Me.tbZugang = New System.Windows.Forms.TextBox
        Me.lblZugang = New System.Windows.Forms.Label
        Me.tbBem = New System.Windows.Forms.TextBox
        Me.lblBem = New System.Windows.Forms.Label
        Me.tbBearb = New System.Windows.Forms.TextBox
        Me.lblBearb = New System.Windows.Forms.Label
        Me.tbDatA = New System.Windows.Forms.TextBox
        Me.lblDatA = New System.Windows.Forms.Label
        Me.tbDatB = New System.Windows.Forms.TextBox
        Me.lblDatB = New System.Windows.Forms.Label
        Me.tbZAStd = New System.Windows.Forms.TextBox
        Me.lblZA = New System.Windows.Forms.Label
        Me.chkExport = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(8, 380)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(104, 23)
        Me.cmdBack.TabIndex = 14
        Me.cmdBack.Text = "<< &Zurück"
        '
        'cmdAbbruch
        '
        Me.cmdAbbruch.Location = New System.Drawing.Point(132, 380)
        Me.cmdAbbruch.Name = "cmdAbbruch"
        Me.cmdAbbruch.Size = New System.Drawing.Size(104, 23)
        Me.cmdAbbruch.TabIndex = 15
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdWeiter
        '
        Me.cmdWeiter.Location = New System.Drawing.Point(488, 380)
        Me.cmdWeiter.Name = "cmdWeiter"
        Me.cmdWeiter.Size = New System.Drawing.Size(104, 23)
        Me.cmdWeiter.TabIndex = 16
        Me.cmdWeiter.Text = "&Weiter >>"
        '
        'tbZugang
        '
        Me.tbZugang.Location = New System.Drawing.Point(136, 196)
        Me.tbZugang.MaxLength = 255
        Me.tbZugang.Name = "tbZugang"
        Me.tbZugang.Size = New System.Drawing.Size(456, 20)
        Me.tbZugang.TabIndex = 11
        Me.tbZugang.Text = ""
        '
        'lblZugang
        '
        Me.lblZugang.Location = New System.Drawing.Point(8, 196)
        Me.lblZugang.Name = "lblZugang"
        Me.lblZugang.Size = New System.Drawing.Size(88, 23)
        Me.lblZugang.TabIndex = 10
        Me.lblZugang.Text = "Zugangsklasse"
        '
        'tbBem
        '
        Me.tbBem.Location = New System.Drawing.Point(136, 160)
        Me.tbBem.MaxLength = 255
        Me.tbBem.Name = "tbBem"
        Me.tbBem.Size = New System.Drawing.Size(456, 20)
        Me.tbBem.TabIndex = 9
        Me.tbBem.Text = ""
        '
        'lblBem
        '
        Me.lblBem.Location = New System.Drawing.Point(8, 160)
        Me.lblBem.Name = "lblBem"
        Me.lblBem.Size = New System.Drawing.Size(88, 23)
        Me.lblBem.TabIndex = 8
        Me.lblBem.Text = "Bemerkungen"
        '
        'tbBearb
        '
        Me.tbBearb.Location = New System.Drawing.Point(136, 124)
        Me.tbBearb.MaxLength = 255
        Me.tbBearb.Name = "tbBearb"
        Me.tbBearb.Size = New System.Drawing.Size(456, 20)
        Me.tbBearb.TabIndex = 7
        Me.tbBearb.Text = ""
        '
        'lblBearb
        '
        Me.lblBearb.Location = New System.Drawing.Point(8, 124)
        Me.lblBearb.Name = "lblBearb"
        Me.lblBearb.Size = New System.Drawing.Size(96, 23)
        Me.lblBearb.TabIndex = 6
        Me.lblBearb.Text = "Bearbeiter im ZA"
        '
        'tbDatA
        '
        Me.tbDatA.Location = New System.Drawing.Point(136, 52)
        Me.tbDatA.MaxLength = 255
        Me.tbDatA.Name = "tbDatA"
        Me.tbDatA.Size = New System.Drawing.Size(456, 20)
        Me.tbDatA.TabIndex = 3
        Me.tbDatA.Text = ""
        '
        'lblDatA
        '
        Me.lblDatA.Location = New System.Drawing.Point(8, 52)
        Me.lblDatA.Name = "lblDatA"
        Me.lblDatA.Size = New System.Drawing.Size(128, 23)
        Me.lblDatA.TabIndex = 2
        Me.lblDatA.Text = "Datum der Archivierung"
        '
        'tbDatB
        '
        Me.tbDatB.Location = New System.Drawing.Point(136, 88)
        Me.tbDatB.MaxLength = 255
        Me.tbDatB.Name = "tbDatB"
        Me.tbDatB.Size = New System.Drawing.Size(456, 20)
        Me.tbDatB.TabIndex = 5
        Me.tbDatB.Text = ""
        '
        'lblDatB
        '
        Me.lblDatB.Location = New System.Drawing.Point(8, 88)
        Me.lblDatB.Name = "lblDatB"
        Me.lblDatB.Size = New System.Drawing.Size(128, 23)
        Me.lblDatB.TabIndex = 4
        Me.lblDatB.Text = "Datum der Bearbeitung"
        '
        'tbZAStd
        '
        Me.tbZAStd.Location = New System.Drawing.Point(136, 16)
        Me.tbZAStd.MaxLength = 255
        Me.tbZAStd.Name = "tbZAStd"
        Me.tbZAStd.Size = New System.Drawing.Size(456, 20)
        Me.tbZAStd.TabIndex = 1
        Me.tbZAStd.Text = ""
        '
        'lblZA
        '
        Me.lblZA.Location = New System.Drawing.Point(8, 16)
        Me.lblZA.Name = "lblZA"
        Me.lblZA.Size = New System.Drawing.Size(120, 23)
        Me.lblZA.TabIndex = 0
        Me.lblZA.Text = "ZA_Studiennummer"
        '
        'chkExport
        '
        Me.chkExport.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkExport.Checked = True
        Me.chkExport.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkExport.Location = New System.Drawing.Point(8, 232)
        Me.chkExport.Name = "chkExport"
        Me.chkExport.Size = New System.Drawing.Size(144, 24)
        Me.chkExport.TabIndex = 13
        Me.chkExport.Text = "exportierbar"
        '
        'Metadaten2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 423)
        Me.Controls.Add(Me.chkExport)
        Me.Controls.Add(Me.tbZAStd)
        Me.Controls.Add(Me.lblZA)
        Me.Controls.Add(Me.tbDatB)
        Me.Controls.Add(Me.lblDatB)
        Me.Controls.Add(Me.tbDatA)
        Me.Controls.Add(Me.lblDatA)
        Me.Controls.Add(Me.tbBearb)
        Me.Controls.Add(Me.lblBearb)
        Me.Controls.Add(Me.tbBem)
        Me.Controls.Add(Me.lblBem)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdWeiter)
        Me.Controls.Add(Me.tbZugang)
        Me.Controls.Add(Me.lblZugang)
        Me.Name = "Metadaten2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Try
            md1.Show()
            Me.Hide()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub cmdWeiter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWeiter.Click
        Try
            ZAStd = tbZAStd.Text
            DatA = tbDatA.Text
            DatB = tbDatB.Text
            Bearb = tbBearb.Text
            Bem = tbBem.Text
            Zugang = tbZugang.Text
            Exportierbar = IIf(chkExport.Checked, 1, 0)

            If IsNothing(md3) Then
                md3 = New Metadaten3
            End If
            Me.Hide()
            md3.Show()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub cmdAbbruch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbbruch.Click
        closeForms()

    End Sub



    Private Sub Metadaten2_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        If Not onAbbr Then
            onAbbr = True
            If Not IsNothing(anm) Then
                anm.Close()
            End If
            If Not IsNothing(md1) Then
                md1.Close()
            End If
            If Not IsNothing(da) Then
                da.Close()
            End If
            sf.Show()
        End If
    End Sub

    Private Sub Metadaten2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.tbZAStd.Text = ZAStd
        Me.tbDatA.Text = DatA
        Me.tbDatB.Text = DatB
        Me.tbBearb.Text = Bearb
        Me.tbBem.Text = Bem
        Me.tbZugang.Text = Zugang
        Me.chkExport.Checked = (Exportierbar = 1)
    End Sub
End Class

