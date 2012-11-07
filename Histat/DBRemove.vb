
Module DBRemove

    Private conn As System.Data.Odbc.OdbcConnection
   Private adapt As OdbcDataAdapter
   Private ds As DataSet

   Public Sub del_Projekt(ByVal id_projekt As String)
      Dim cmd As OdbcCommand
      Try
            conn = New System.Data.Odbc.OdbcConnection(conStr)
         ds = New DataSet()
         conn.Open()
         del_Schluesselmaske(id_projekt)
         cmd = New OdbcCommand("DELETE FROM Aka_Projekte WHERE ID_Projekt = '" & id_projekt & "';", conn)
         cmd.ExecuteNonQuery()
         conn.Close()
      Catch ex As Exception
         Throw ex
      Finally
         ds = Nothing
         adapt = Nothing
         conn = Nothing
      End Try

   End Sub

   Public Sub del_Daten(ByVal id_projekt As String)
      Dim cmd As OdbcCommand
      Try
            conn = New System.Data.Odbc.OdbcConnection(conStr)
         ds = New DataSet()
         conn.Open()
         del_Schluesselmaske(id_projekt)
         cmd = New OdbcCommand("UPDATE Aka_Projekte SET Anzahl_Zeitreihen = 0 WHERE ID_Projekt = '" & id_projekt & "';", conn)
         cmd.ExecuteNonQuery()
         conn.Close()
      Catch ex As Exception
         Throw ex
      Finally
         ds = Nothing
         adapt = Nothing
         conn = Nothing
      End Try

   End Sub

   Private Sub del_Schluesselmaske(ByVal id_projekt As String)
        Dim dt As System.Data.DataTable
      Dim dr As DataRow
        Dim param_LZ As OdbcParameter
      Dim param_DA As OdbcParameter
      Dim cmd As OdbcCommand
      Dim cmd_LZ As OdbcCommand
      Dim cmd_DA As OdbcCommand

      Try
         adapt = New OdbcDataAdapter("SELECT * FROM Aka_Schluesselmaske WHERE ID_Projekt = '" & id_projekt & "';", conn)
         adapt.Fill(ds, "DS_Aka_Schluesselmaske")

         cmd = New OdbcCommand("DELETE FROM Aka_Schluesselmaske WHERE ID_Projekt = '" & id_projekt & "';", conn)

         cmd_DA = New OdbcCommand("DELETE FROM Daten__Aka WHERE ID_HS = ?;", conn)
         param_DA = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd_DA.Parameters.Add(param_DA)

         cmd_LZ = New OdbcCommand("DELETE FROM Lit_ZR WHERE ID_HS = ?;", conn)
         param_LZ = New OdbcParameter("@ID_HS", OdbcType.VarChar, 32, "ID_HS")
         cmd_LZ.Parameters.Add(param_LZ)

         dt = ds.Tables("DS_Aka_Schluesselmaske")
         For Each dr In dt.Rows
            del_SchluesselCode(dr.Item("ID_HS"))
            param_DA.Value = dr.Item("ID_HS")
            cmd_DA.ExecuteNonQuery()
            param_LZ.Value = dr.Item("ID_HS")
            cmd_LZ.ExecuteNonQuery()
         Next

         cmd.ExecuteNonQuery()
         dt.Clear()
      Catch ex As Exception
         Throw ex
      End Try

   End Sub


   Private Sub del_SchluesselCode(ByVal id_hs As String)
        Dim dt As System.Data.DataTable
      Dim dr As DataRow
      Dim cmd As OdbcCommand
      Try
         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_SchluesselCode WHERE ID_HS = '" & id_hs & "';"
         adapt.Fill(ds, "DS_Aka_SchluesselCode")

         dt = ds.Tables("DS_Aka_SchluesselCode")
         For Each dr In dt.Rows
            del_Codes(dr.Item("ID_CodeKuerz"))
         Next

         cmd = New OdbcCommand("DELETE FROM Aka_SchluesselCode WHERE ID_HS = '" & id_hs & "';", conn)
         cmd.ExecuteNonQuery()
         dt.Clear()
      Catch ex As Exception
         Throw ex
      End Try

   End Sub


   Private Sub del_Codes(ByVal id_CodeKuerz As String)
        Dim dt As System.Data.DataTable
      Dim dr As DataRow
      Dim param_CI As OdbcParameter
      Dim cmd As OdbcCommand
      Dim cmd_CI As OdbcCommand

      Try
         adapt.SelectCommand.CommandText = "SELECT * FROM Aka_Codes WHERE ID_CodeKuerz = '" & id_CodeKuerz & "';"
         adapt.Fill(ds, "DS_Aka_Codes")
         dt = ds.Tables("DS_Aka_Codes")

         cmd_CI = New OdbcCommand("DELETE FROM Aka_CodeInhalt WHERE ID_CodeKuerz = ?;", conn)
         param_CI = New OdbcParameter("@ID_CodeKuerz", OdbcType.VarChar, 32, "ID_CodeKuerz")
         cmd_CI.Parameters.Add(param_CI)

         For Each dr In dt.Rows
            param_CI.Value = dr.Item("ID_CodeKuerz")
            cmd_CI.ExecuteNonQuery()
         Next

         cmd = New OdbcCommand("DELETE FROM Aka_Codes WHERE ID_CodeKuerz = '" & id_CodeKuerz & "';", conn)
         cmd.ExecuteNonQuery()

         dt.Clear()
      Catch ex As Exception
         Throw ex
      End Try

   End Sub

End Module
