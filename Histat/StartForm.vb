Imports System.IO

Public Class StartForm
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbThema As System.Windows.Forms.ComboBox
   Friend WithEvents Button3 As System.Windows.Forms.Button
   Friend WithEvents Button4 As System.Windows.Forms.Button
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(StartForm))
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.cbThema = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(80, Byte), Integer), CType(CType(161, Byte), Integer))
        Me.Button1.Location = New System.Drawing.Point(174, 144)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(272, 27)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "&Neues Projekt anlegen"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(174, 273)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(272, 27)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Vorhandenes Projekt &löschen"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(427, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(149, 44)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(35, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(313, 58)
        Me.PictureBox2.TabIndex = 3
        Me.PictureBox2.TabStop = False
        '
        'cbThema
        '
        Me.cbThema.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbThema.Location = New System.Drawing.Point(88, 104)
        Me.cbThema.Name = "cbThema"
        Me.cbThema.Size = New System.Drawing.Size(488, 21)
        Me.cbThema.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(32, 106)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Themen:"
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(80, Byte), Integer), CType(CType(161, Byte), Integer))
        Me.Button3.Location = New System.Drawing.Point(174, 230)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(272, 27)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = " Projekt&daten neu importieren"
        '
        'Button4
        '
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(80, Byte), Integer), CType(CType(161, Byte), Integer))
        Me.Button4.Location = New System.Drawing.Point(174, 187)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(272, 27)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Projektbeschreibung &bearbeiten"
        '
        'StartForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 423)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbThema)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.ForeColor = System.Drawing.Color.DarkRed
        Me.Name = "StartForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Try
            If (cbThema.SelectedIndex = -1) Then
                MsgBox("Bitte wählen Sie ein Thema aus!", , "Histat")
                cbThema.Focus()
                Exit Sub
            End If
            Me.Hide()
            status = Globals.tStatus.onImport
            anm = New Anmeldung()
            anm.Show()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If (cbThema.SelectedIndex = -1) Then
                MsgBox("Bitte wählen Sie ein Thema aus!", , "Histat")
                cbThema.Focus()
                Exit Sub
            End If
            status = Globals.tStatus.onDelete
            Me.Hide()
            pa = New Projektauswahl()
            pa.Show()

        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub StartForm_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            HistatLog.Flush()
            HistatLog.Close()
            HistatLog = Nothing
        Catch ex As Exception
        End Try
        System.Windows.Forms.Application.Exit()
    End Sub


    Private Sub StartForm_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            anm = Nothing
            md1 = Nothing
            md2 = Nothing
            md3 = Nothing
            da = Nothing
            pa = Nothing
            onAbbr = False
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try

    End Sub

    Private Sub StartForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

            sSql = "SELECT ID_Thema, Thema, Position FROM Aka_Themen ORDER BY Position;"
            adapt = New OdbcDataAdapter(sSql, conn)
            adapt.Fill(ds, "DS_Aka_Themen")
            dt = ds.Tables("DS_Aka_Themen")
            key(0) = dt.Columns(0)
            dt.PrimaryKey = key
            conn.Close()

            With cbThema
                .DataSource = dt
                .DisplayMember = "Thema"
                .ValueMember = "ID_Thema"
                .SelectedIndex = -1
            End With
            adapt = Nothing
            ds = Nothing
            conn = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub cbThema_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbThema.SelectedIndexChanged
        Dim dr As DataRow
        Try
            With cbThema
                If Not IsNothing(.SelectedItem) Then
                    dr = .SelectedItem.row
                    ID_Thema = dr.Item("ID_Thema")
                End If
            End With
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If (cbThema.SelectedIndex = -1) Then
                MsgBox("Bitte wählen Sie ein Thema aus!", , "Histat")
                cbThema.Focus()
                Exit Sub
            End If
            status = Globals.tStatus.onDataImport
            Me.Hide()
            pa = New Projektauswahl()
            pa.Show()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If (cbThema.SelectedIndex = -1) Then
                MsgBox("Bitte wählen Sie ein Thema aus!", , "Histat")
                cbThema.Focus()
                Exit Sub
            End If
            Me.Hide()
            status = Globals.tStatus.onMetadataUpdate
            pa = New Projektauswahl()
            pa.Show()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub
End Class
