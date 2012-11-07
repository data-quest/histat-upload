Imports System.io

Public Class Dateiauswahl
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
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents LB As System.Windows.Forms.ListBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdOK As System.Windows.Forms.Button
    Friend WithEvents cmdAbbruch As System.Windows.Forms.Button
    Friend WithEvents cmdBack As System.Windows.Forms.Button
    Friend WithEvents SBar As System.Windows.Forms.StatusBar
    Friend WithEvents SPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents SPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents OFD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.LB = New System.Windows.Forms.ListBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.OFD = New System.Windows.Forms.OpenFileDialog
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.cmdAbbruch = New System.Windows.Forms.Button
        Me.cmdBack = New System.Windows.Forms.Button
        Me.SBar = New System.Windows.Forms.StatusBar
        Me.SPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.SPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.SPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LB
        '
        Me.LB.HorizontalScrollbar = True
        Me.LB.Location = New System.Drawing.Point(16, 48)
        Me.LB.Name = "LB"
        Me.LB.Size = New System.Drawing.Size(576, 225)
        Me.LB.TabIndex = 0
        Me.LB.TabStop = False
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(16, 331)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(104, 23)
        Me.cmdAdd.TabIndex = 1
        Me.cmdAdd.Text = "&Hinzufügen"
        '
        'OFD
        '
        Me.OFD.Filter = "Excel-Dateien|*.xls"
        Me.OFD.Multiselect = True
        '
        'cmdRemove
        '
        Me.cmdRemove.AccessibleDescription = ""
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.Location = New System.Drawing.Point(132, 331)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(104, 23)
        Me.cmdRemove.TabIndex = 2
        Me.cmdRemove.Text = "&Entfernen"
        '
        'cmdOK
        '
        Me.cmdOK.Enabled = False
        Me.cmdOK.Location = New System.Drawing.Point(488, 363)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(104, 23)
        Me.cmdOK.TabIndex = 5
        Me.cmdOK.Text = "&Fertigstellen"
        '
        'cmdAbbruch
        '
        Me.cmdAbbruch.Location = New System.Drawing.Point(132, 363)
        Me.cmdAbbruch.Name = "cmdAbbruch"
        Me.cmdAbbruch.Size = New System.Drawing.Size(104, 23)
        Me.cmdAbbruch.TabIndex = 4
        Me.cmdAbbruch.Text = "&Abbrechen"
        '
        'cmdBack
        '
        Me.cmdBack.Location = New System.Drawing.Point(16, 363)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(104, 23)
        Me.cmdBack.TabIndex = 3
        Me.cmdBack.Text = "<< &Zurück"
        '
        'SBar
        '
        Me.SBar.Location = New System.Drawing.Point(0, 401)
        Me.SBar.Name = "SBar"
        Me.SBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.SPanel1, Me.SPanel2})
        Me.SBar.ShowPanels = True
        Me.SBar.Size = New System.Drawing.Size(612, 22)
        Me.SBar.TabIndex = 0
        '
        'SPanel1
        '
        Me.SPanel1.Name = "SPanel1"
        Me.SPanel1.Width = 240
        '
        'SPanel2
        '
        Me.SPanel2.Name = "SPanel2"
        Me.SPanel2.Width = 362
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(320, 24)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Folgende Excel-Dateien werden importiert:"
        '
        'Dateiauswahl
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(612, 423)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SBar)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdAbbruch)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.LB)
        Me.Name = "Dateiauswahl"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Histat-Import"
        CType(Me.SPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim dr As DialogResult
        Dim i As Int16
        Dim fname As String
        Try
            With OFD
                dr = .ShowDialog()
                If dr = DialogResult.OK Then
                    For i = 0 To .FileNames().GetUpperBound(0)
                        fname = .FileNames(i)
                        If Not LB.Items.Contains(fname) Then
                            LB.Items.Add(fname)
                        End If
                    Next
                End If
            End With
            If LB.Items.Count > 0 Then
                cmdRemove.Enabled = True
                cmdOK.Enabled = True
            End If
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Try
            With LB
                .Items.Remove(.SelectedItem)
                If .Items.Count = 0 Then
                    cmdRemove.Enabled = False
                    cmdOK.Enabled = False
                End If
            End With
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        Dim i As Int32
        Dim ub As Int32
        Dim azr As Int32
        Dim zr As String

        Try
            zr = ""
            cmdAdd.Enabled = False
            cmdRemove.Enabled = False
            cmdOK.Enabled = False
            cmdAbbruch.Enabled = False
            cmdBack.Enabled = False

            ub = LB.Items.Count - 1
            ReDim xlsArr(ub)
            For i = 0 To ub
                xlsArr(i) = LB.Items(i)
            Next

            Me.Cursor = Cursors.WaitCursor

            If checkXlTabs(azr, zr) Then                   ' Import oder Data-Import
                If status = Globals.tStatus.onImport Then
                    ID_Projekt = DBImport.ImportProject(azr, zr)
                Else
                    DBRemove.del_Daten(ID_Projekt)
                End If
                importDaten(ID_Projekt, azr, zr)
            Else
                If status = Globals.tStatus.onImport Then
                    If MsgBox("Wegen einer fehlerhaften Excel-Tabelle wurde kein Datenimport durchgeführt!" & ControlChars.CrLf & _
                          "Sollen die eingegebenen Metadaten trotzdem gespeichert werden?", MsgBoxStyle.YesNo, "Histat-Import") = MsgBoxResult.Yes Then
                        DBImport.ImportProject(azr, zr)
                    End If
                Else
                    MsgBox("Wegen einer fehlerhaften Excel-Tabelle wurde kein Datenimport durchgeführt!", , "Histat-Import")
                End If
                da.SPanel1.Text = "Abbruch"
                openLog()
            End If

            Me.Cursor = Cursors.Default

            da.SPanel2.Text = ""
            da.SBar.Refresh()

        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        Finally
            closeForms()
        End Try

    End Sub

    Private Sub cmdAbbruch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAbbruch.Click
        closeForms()

    End Sub


    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
        Try
            If status = Globals.tStatus.onImport Then
                md3.Show()
            Else
                pa.Show()
            End If
            Me.Hide()
        Catch ex As Exception
            HistatLog.WriteLine(Now & " " & (ex).GetType.Name & ControlChars.CrLf & ex.Message)
            closeForms()
            MsgBox((ex).GetType.Name & ControlChars.CrLf & ex.Message, MsgBoxStyle.Critical, "Histat")
        End Try
    End Sub


    Private Sub Dateiauswahl_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        If Not onAbbr Then
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
            sf.Show()
        End If
    End Sub

    Private Sub importDaten(ByVal strID_Project As String, ByVal azr As Int32, ByVal zr As String)
        Dim exc As Exception
        Dim i As Int32
        Dim l As Int32
        Dim ub As Int32
        Dim xlFile As String

        ub = xlsArr.GetUpperBound(0)
        For i = 0 To ub
            xlFile = xlsArr(i)
            l = InStrRev(xlFile, "\")
            SPanel1.Text = Mid(xlFile, l + 1)
            exc = New Exception("neu")
            If Not DBImport.ImportData(xlFile, exc, strID_Project, azr, zr) Then
                da.SPanel1.Text = "Abbruch: " & Mid(xlFile, l + 1)
                MsgBox(exc.Message, MsgBoxStyle.Critical, "Histat-Import")
                HistatLog.WriteLine(Now & " " & (exc).GetType.Name & ControlChars.CrLf & exc.Message)
                HistatLog.WriteLine(Now & " Der Datenimport ist fehlgeschlagen." & vbCrLf)
                If status = Globals.tStatus.onImport Then
                    If MsgBox("Der Datenimport ist fehlgeschlagen!" & ControlChars.CrLf & _
                       "Sollen die eingegebenen Metadaten trotzdem gespeichert werden?", MsgBoxStyle.YesNo, "Histat-Import") = MsgBoxResult.No Then
                        DBRemove.del_Projekt(strID_Project)
                    Else
                        DBRemove.del_Daten(strID_Project)
                    End If
                Else
                    DBRemove.del_Daten(strID_Project)
                    MsgBox("Der Datenimport ist fehlgeschlagen!" & ControlChars.CrLf & _
                       "Die alten Zeitreihen des Projekts wurden gelöscht," & ControlChars.CrLf & _
                       "d.h. nur noch die Metadaten sind gespeichert!", , "Histat-Import")
                End If

                DBImport.dispose()
                Exit Sub
            Else
                DBImport.dispose()
            End If
        Next

        If azr > 0 Then
            MsgBox("Der Histat-Import wurde erfolgreich abgeschlossen!" & ControlChars.CrLf & _
            "Für den Zeitraum " & zr & " wurden " & azr & " Zeitreihen erstellt!", , "Histat-Import")
            HistatLog.WriteLine(Now & " Der Histat-Import wurde erfolgreich abgeschlossen." & vbCrLf)
        Else
            If MsgBox("Der Histat-Import wurde abgeschlossen!" & ControlChars.CrLf & _
                "Aber für den Zeitraum " & zr & " wurden keine Zeitreihen erstellt!" & ControlChars.CrLf & _
                "Sollen die eingegebenen Metadaten trotzdem gespeichert werden?", MsgBoxStyle.YesNo, "Histat-Import") = MsgBoxResult.No Then
                DBRemove.del_Projekt(strID_Project)
                HistatLog.WriteLine(Now & " Der Histat-Import der Metadaten wurde abgebrochen." & vbCrLf)
            Else
                HistatLog.WriteLine(Now & " Der Histat-Import wurde ohne Datenimport abgeschlossen." & vbCrLf)
            End If
        End If
        DBImport.dispose()
        da.SPanel1.Text = "Fertig"
    End Sub


End Class
