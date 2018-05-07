Imports System.Data.SqlClient
Public Class BaseDatos
#Region "Propiedades"
    Public CadenaConexion As String
    Private Transacciones As ArrayList
    Private NC1 As New SqlConnection
    Private Trans As SqlTransaction
    Private TransOpen As Boolean
#End Region

#Region "Constructor y Conexiones"

    Public Sub New(ByVal CadenaConexion As String, Optional pvb_ActualizarConexion As Boolean = True)
        'Me.CadenaConexion = CadenaConexion

        NC1.ConnectionString = CadenaConexion

        If pvb_ActualizarConexion Then
            Try
                Dim vlo_Datos As New DataTable("Tmp")

                vlo_Datos.Columns.Add("Tiempo")

                vlo_Datos.Rows.Add(DateTime.Now.ToString("yyyyMMdd hh:mm:ss tt"))

                vlo_Datos.WriteXml("Tiempo.xml")
            Catch
            End Try
        End If
    End Sub

    Sub New()
        ' TODO: Complete member initialization 
    End Sub

    Public Function ConectarDB() As Boolean
        Try
            'NC1.ConnectionString = CadenaConexion
            NC1.Open()
        Catch ex As Exception
            'MsgBox("NO se puede iniciar conectar, error 1015: " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function DesconectarDB() As Boolean
        Try
            NC1.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function OpenTrans() As Boolean
        Try
            'NC1.ConnectionString = CadenaConexion
            If NC1.State <> ConnectionState.Open Then
                NC1.Open()
            End If
            Trans = NC1.BeginTransaction
            TransOpen = True
        Catch ex As Exception
            TransOpen = False
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function CloseTrans() As Boolean
        Try
            Trans.Commit()
            NC1.Close()
        Catch ex As Exception
            TransOpen = False
            MsgBox(ex.Message)
            Return False
        End Try

        TransOpen = False

        Return True
    End Function

    Public Sub RollBackTrans()
        Try
            Trans.Rollback()
            TransOpen = False
            'NC1.Dispose()
        Catch ex As Exception
        End Try

        TransOpen = False
    End Sub
#End Region

#Region "Metodos Ejecuta SQL"
    Public Function EjecutarSQL(ByVal CMD As SqlCommand, Optional ByVal pvc_Serializable As Boolean = False) As DLLTipos.GENTipos.Terror
        If TransOpen Then
            Return EjecutarSQL_T(CMD, pvc_Serializable)
        Else
            Return EjecutarSQL_N(CMD, pvc_Serializable)
        End If
    End Function

    Private Function EjecutarSQL_N(ByVal CMD As SqlCommand, ByVal pvc_Serializable As Boolean) As DLLTipos.GENTipos.Terror
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""
        Try

            If Not ConectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description

                Return control
            End If

            If pvc_Serializable Then
                CMD.CommandType = CommandType.StoredProcedure
            End If

            CMD.Connection = NC1
            CMD.ExecuteNonQuery()

            If Not DesconectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
            End If
        Catch ex As Exception
            NC1.Close()
            control.estado = False
            control.nerror = Err.Number
            control.mensaje = Err.Description
        End Try
        Return control
    End Function


    Private Function EjecutarSQL_T(ByVal CMD As SqlCommand, ByVal pvc_Serializable As Boolean) As DLLTipos.GENTipos.Terror
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""

        Try
            If pvc_Serializable Then
                CMD.CommandType = CommandType.StoredProcedure
            End If

            CMD.Connection = NC1
            CMD.Transaction = Trans
            CMD.ExecuteNonQuery()
        Catch ex As Exception
            Try
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
                RollBackTrans()

            Catch
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
                'no hacer nada 
            End Try
        End Try
        CMD.Dispose()
        Return control
    End Function

    Public Function EjecutaSQL(ByVal pvo_Comando As SqlCommand, Optional ByVal pvb_StoreProcedure As Boolean = False) As DLLTipos.GENTipos.Terror

        If TransOpen Then
            Return EjecutaSQL_T(pvo_Comando, pvb_StoreProcedure)
        Else
            Return EjecutaSQL_N(pvo_Comando, pvb_StoreProcedure)
        End If

    End Function

    Private Function EjecutaSQL_T(ByVal pvo_Comando As SqlCommand, ByVal pvb_StoreProcedure As Boolean) As DLLTipos.GENTipos.Terror
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""

        Try
            pvo_Comando.Connection = NC1
            pvo_Comando.Transaction = Trans

            If (pvb_StoreProcedure) Then
                pvo_Comando.CommandType = CommandType.StoredProcedure
            End If

            pvo_Comando.ExecuteNonQuery()
        Catch vlo_Error As Exception
            RollBackTrans()
            control.estado = False
            control.nerror = Err.Number
            control.mensaje = Err.Description
        End Try

        Return control
    End Function

    Private Function EjecutaSQL_N(ByVal pvo_Comando As SqlCommand, ByVal pvb_StoreProcedure As Boolean) As DLLTipos.GENTipos.Terror
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""

        Try
            If Not ConectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description

                Return control
            End If

            If (pvb_StoreProcedure) Then
                pvo_Comando.CommandType = CommandType.StoredProcedure
            End If

            pvo_Comando.Connection = NC1

            pvo_Comando.ExecuteNonQuery()

            DesconectarDB()
        Catch ex As Exception
            control.estado = False
            control.nerror = Err.Number
            control.mensaje = Err.Description
        End Try

        Return control
    End Function


    Public Function EjecutaSQL(ByVal SQL As String, Optional ByVal getconsecutivo As Boolean = False) As DLLTipos.GENTipos.Terror
        If TransOpen Then
            Return EjecutaSQL_T(SQL, getconsecutivo)
        Else
            Return EjecutaSQL_N(SQL, getconsecutivo)
        End If
    End Function

    Private Function EjecutaSQL_N(ByVal SQL As String, ByVal getconsecutivo As Boolean) As DLLTipos.GENTipos.Terror
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""
        Try

            If Not ConectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
            End If

            Dim Comando As New SqlCommand(SQL, NC1)
            'If getconsecutivo = False Then
            Comando.ExecuteNonQuery()
            'Else
            'control.tag = Comando.ExecuteScalar()
            'End If


            If Not DesconectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
            End If
        Catch ex As Exception
            NC1.Close()
            control.estado = False
            control.nerror = Err.Number
            control.mensaje = Err.Description
        End Try
        Return control
    End Function

    Private Function EjecutaSQL_T(ByVal sql As String, ByVal getconsecutivo As Boolean) As DLLTipos.GENTipos.Terror
        Dim Comando As New SqlCommand(sql, NC1)
        Dim control As New DLLTipos.GENTipos.Terror
        control.estado = True
        control.nerror = 0
        control.mensaje = ""
        Try
            Comando.Connection = NC1
            Comando.Transaction = Trans
            If getconsecutivo = False Then
                Comando.ExecuteNonQuery()

            Else
                control.tag = Comando.ExecuteScalar.ToString
            End If
        Catch ex As Exception
            Try
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
                RollBackTrans()

            Catch
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description
                'no hacer nada 
            End Try
        End Try
        Comando.Dispose()
        Return control
    End Function
#End Region

#Region "Metodos Data Set"
    Public Function LlenarDS(ByRef DS As DataSet, ByVal CMD As SqlCommand, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        If TransOpen Then 'Hay una Transaccion Abierta
            Return LlenarDS_T(DS, CMD, Nombre)
        Else
            Return LlenarDS_N(DS, CMD, Nombre)
        End If
    End Function

    Private Function LlenarDS_N(ByRef DS As DataSet, ByVal CMD As SqlCommand, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        Try
            If Not ConectarDB() Then Return False
            CMD.Connection = NC1

            If Nombre = "DatosStore" Then
                CMD.CommandType = CommandType.StoredProcedure
            End If

            CMD.CommandTimeout = 0

            Dim Adaptador As New SqlDataAdapter(CMD)
            Adaptador.Fill(DS, Nombre)
            If Not DesconectarDB() Then Return False
        Catch ex As Exception
            DesconectarDB()
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Function LlenarDS_T(ByRef DS As DataSet, ByVal CMD As SqlCommand, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        Try
            CMD.Connection = NC1
            CMD.Transaction = Trans

            If Nombre = "DatosStore" Then
                CMD.CommandType = CommandType.StoredProcedure
            End If

            CMD.CommandTimeout = 0

            CMD.ExecuteNonQuery()
            Dim adaptador As New SqlDataAdapter(CMD)
            adaptador.Fill(DS, Nombre)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function LlenarDS(ByRef DS As DataSet, ByVal SQL As String, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        If TransOpen Then 'Hay una Transaccion Abierta
            Return LlenarDS_T(DS, SQL, Nombre)
        Else
            Return LlenarDS_N(DS, SQL, Nombre)
        End If
    End Function

    Private Function LlenarDS_T(ByRef DS As DataSet, ByVal SQL As String, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        Try
            Dim CommandoSQL As New SqlCommand(SQL, NC1, Trans)

            CommandoSQL.CommandTimeout = 0
            CommandoSQL.ExecuteNonQuery()
            Dim adaptador As New SqlDataAdapter(CommandoSQL)
            adaptador.Fill(DS, Nombre)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Private Function LlenarDS_N(ByRef DS As DataSet, ByVal SQL As String, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        Try
            If Not ConectarDB() Then Return False

            Dim CommandoSQL As New SqlCommand(SQL, NC1, Trans)

            CommandoSQL.CommandTimeout = 0
            CommandoSQL.ExecuteNonQuery()
            Dim adaptador As New SqlDataAdapter(CommandoSQL)
            adaptador.Fill(DS, Nombre)

            'Dim Adaptador As New SqlDataAdapter(SQL, NC1)



            'Adaptador.Fill(DS, Nombre)
            If Not DesconectarDB() Then Return False
        Catch ex As Exception
            DesconectarDB()
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function LlenarDR(ByRef DR As System.Data.SqlClient.SqlDataReader, ByVal SQL As String, Optional ByVal Nombre As String = "NoDefinido") As Boolean
        Try
            Dim Comando As New SqlClient.SqlCommand(SQL)
            Comando.Connection = NC1
            DR = Comando.ExecuteReader
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
        Return True
    End Function
    Public Function LLenaDataSet(ByVal pSQL As String, Optional ByVal pNombre As String = "NoDefinido") As DataSet
        Dim DS As New DataSet

        LlenarDS(DS, pSQL, pNombre)

        Return DS

        'Dim Query As String = pSQL
        'Dim Command As New SqlCommand(Query)
        'Dim Reader As SqlDataReader
        'Dim dt As New DataTable
        'Dim ds As New DataSet

        'dt.TableName = pNombre

        'Try
        '    If TransOpen Then 'Hay una Transaccion Abierta
        '        Command.Connection = NC1
        '        Command.Transaction = Trans
        '        Command.CommandTimeout = 0
        '        Reader = Command.ExecuteReader
        '        dt.Load(Reader, LoadOption.OverwriteChanges)
        '        ds.Tables.Add(dt)
        '    Else
        '        If ConectarDB() Then
        '            Command.Connection = NC1
        '            Command.CommandTimeout = 0
        '            Reader = Command.ExecuteReader
        '            dt.Load(Reader, LoadOption.OverwriteChanges)
        '            ds.Tables.Add(dt)
        '        End If
        '    End If

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        'If TransOpen = False Then NC1.Close()

        'Return ds

    End Function

    ''' <summary>
    ''' Carga un dataset
    ''' </summary>
    ''' <param name="pvo_Comando">Comando con los parametros y Intruccion Sql</param>
    ''' <param name="pvc_Nombre">Nombre de la tabla que se devuelve</param>
    ''' <param name="pvb_StoreProcedure">Indica si es un storeprocedure o un query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LlenarDataSet(ByVal pvo_Comando As SqlCommand, ByVal pvc_Nombre As String, Optional ByVal pvb_StoreProcedure As Boolean = False) As DLLTipos.GENTipos.Terror
        'verifica que si la trasaccion esta abierta
        If (TransOpen) Then
            'llena el data set en una transaccion
            Return LlenarDS_Transaccional(pvo_Comando, pvc_Nombre, pvb_StoreProcedure)
        Else
            Return LlenarDS_NoTransaccional(pvo_Comando, pvc_Nombre, pvb_StoreProcedure)
        End If
    End Function

    ''' <summary>
    ''' Carga un dataset en transaccion
    ''' </summary>
    ''' <param name="pvo_Comando">Comando con los parametros y Intruccion Sql</param>
    ''' <param name="pvc_Nombre">Nombre de la tabla que se devuelve</param>
    ''' <param name="pvb_StoreProcedure">Indica si es un storeprocedure o un query</param>
    ''' <returns>retorna un arreglo con el dataset que se desea</returns>
    ''' <remarks></remarks>
    Private Function LlenarDS_Transaccional(ByVal pvo_Comando As SqlCommand, ByVal pvc_Nombre As String, ByVal pvb_StoreProcedure As Boolean) As DLLTipos.GENTipos.Terror
        Dim vlo_Datos As New DataSet()
        Dim vlo_Adaptador As New SqlDataAdapter
        Dim control As New DLLTipos.GENTipos.Terror

        control.estado = True
        control.nerror = 0
        control.mensaje = ""

        Try
            'ejecuta el query
            pvo_Comando.Connection = NC1
            pvo_Comando.Transaction = Trans

            If (pvb_StoreProcedure) Then
                pvo_Comando.CommandType = CommandType.StoredProcedure
            End If

            pvo_Comando.ExecuteNonQuery()

            'llena el dataset
            vlo_Adaptador = New SqlDataAdapter(pvo_Comando)
            vlo_Adaptador.Fill(vlo_Datos, pvc_Nombre)
        Catch ex As Exception
            RollBackTrans()

            control.estado = False
            control.mensaje = ex.Message
        End Try

        control.vlo_DataSet = vlo_Datos

        Return control
    End Function

    ''' <summary>
    ''' Carga un dataset sin transaccion
    ''' </summary>
    ''' <param name="pvo_Comando">Comando con los parametros y Intruccion Sql</param>
    ''' <param name="pvc_Nombre">Nombre de la tabla que se devuelve</param>
    ''' <param name="pvb_StoreProcedure">Indica si es un storeprocedure o un query</param>
    ''' <returns>retorna un arreglo con el dataset que se desea</returns>
    ''' <remarks></remarks>
    Private Function LlenarDS_NoTransaccional(ByVal pvo_Comando As SqlCommand, ByVal pvc_Nombre As String, ByVal pvb_StoreProcedure As Boolean) As DLLTipos.GENTipos.Terror
        Dim vlo_Adaptador As New SqlDataAdapter
        Dim vlo_Datos As New DataSet()
        Dim control As New DLLTipos.GENTipos.Terror

        control.estado = True
        control.nerror = 0
        control.mensaje = ""

        Try
            If Not ConectarDB() Then
                control.estado = False
                control.nerror = Err.Number
                control.mensaje = Err.Description

                Return control
            End If

            pvo_Comando.Connection = NC1

            If (pvb_StoreProcedure) Then
                pvo_Comando.CommandType = CommandType.StoredProcedure
            End If

            vlo_Adaptador = New SqlDataAdapter(pvo_Comando)

            vlo_Adaptador.Fill(vlo_Datos, pvc_Nombre)

            DesconectarDB()
        Catch ex As Exception
            DesconectarDB()

            control.estado = False
            control.mensaje = ex.Message
        End Try

        control.vlo_DataSet = vlo_Datos

        Return control
    End Function

    Public Function LlenaDataTable(ByVal pSQL As String, Optional ByVal pNombre As String = "NoDefinido") As DataTable
        Try
            Dim DS As New DataSet

            If LlenarDS(DS, pSQL, pNombre) = False Then
                Dim vlo_tabla As New DataTable(pNombre)

                DS.Tables.Add(vlo_tabla)
            End If

            Return DS.Tables(pNombre).Copy
        Catch ex As Exception
            MsgBox(ex.Message)
            RollBackTrans()
            Return Nothing
        End Try



      

        'Dim Query As String = pSQL
        'Dim Reader As SqlDataReader
        'Dim dt As New DataTable

        'Try
        '    Dim Command As New SqlCommand(Query)

        '    dt.TableName = pNombre

        '    If TransOpen Then 'Hay una Transaccion Abierta
        '        Command.Connection = NC1
        '        Command.Transaction = Trans
        '        Command.CommandTimeout = 0
        '        Reader = Command.ExecuteReader
        '        dt.Load(Reader, LoadOption.OverwriteChanges)
        '    Else
        '        If ConectarDB() Then
        '            Command.Connection = NC1
        '            Command.CommandTimeout = 0
        '            Reader = Command.ExecuteReader
        '            dt.Load(Reader, LoadOption.OverwriteChanges)
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        'If TransOpen = False Then NC1.Close()

        'Return dt
    End Function
#End Region

#Region "Metodos Otros"
    Public Function DatoRepetido(ByVal Tabla As String, ByVal campo As String, ByVal Valor As String, ByVal tipo As String, ByVal condicion As String, Optional ByVal Ind_Trans As Boolean = False) As Boolean
        Dim SQL, aux As String
        Dim DS As New DataSet
        aux = IIf(UCase(Trim(tipo)) <> "N", "'", "")
        If condicion = "" Then
            SQL = "select * from " & Tabla & " where " & campo & " = " & aux & Valor & aux & " "
        Else
            SQL = "select * from " & Tabla & " where " & condicion
        End If
        If Ind_Trans Then
            LlenarDS(DS, SQL)
        Else
            LlenarDS_T(DS, SQL)
        End If
        If DS.Tables(0).Rows.Count > 0 Then
            Return True
        End If
    End Function
#End Region


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
