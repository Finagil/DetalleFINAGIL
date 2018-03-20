Option Explicit On

Imports System.Data.SqlClient
Imports System.Math

Public Class frmLlenado

    ' Declaración de variables de conexión ADO. NET de alcance privado

    Dim dtTIIE As New DataTable()
    Dim drTIIE As DataRow
    Dim myKeySearch(0) As String

    ' Genero la tabla que contiene las TIIE promedio por mes 
    ' Para FINAGIL considera todos los días del mes y redondea a 4 decimales

    Private Sub frmLlenado_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dtTIIE = TIIEavg("FINAGIL")
        btnProcesar_Click(Nothing, Nothing)
        'Button1_Click(Nothing, Nothing)
        End
    End Sub


    Private Sub btnProcesar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcesar.Click
        Dim ERRR As New System.IO.StreamWriter("c:\Files\Errores.txt", System.IO.FileMode.Append, System.Text.Encoding.GetEncoding(1252))

        ' Declaración de variables de conexión ADO .NET
        Dim cnAgil As New SqlConnection("Server=SERVER-RAID; DataBase=Production; User ID = 'User_PRO'; pwd = 'User_PRO2015'")
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim daMinistracion As New SqlDataAdapter(cm1)
        Dim daDetalle As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet()
        Dim drMinistracion As DataRow
        Dim drDetalle As DataRow

        Dim strInsert As String

        ' Declaración de variables de datos

        Dim cAnexo As String = ""
        Dim cCiclo As String = ""
        Dim cCliente As String = ""
        Dim cDocumento As String = ""
        Dim cFechaFinal As String = ""
        Dim cFechaInicial As String = ""
        Dim cFechaPago As String = ""
        Dim cTipta As String = ""
        Dim cFondeo As String = ""
        Dim nConsecutivo As Integer = 0
        Dim nDias As Integer = 0
        Dim nDiferencial As Decimal = 0
        Dim nFEGA As Decimal = 0
        Dim nGarantia As Decimal = 0
        Dim nImporte As Decimal = 0
        Dim nIntereses As Decimal = 0
        Dim nSaldoFinal As Decimal = 0
        Dim nSaldoInicial As Decimal = 0
        Dim nTasa As Decimal = 0
        Dim nTasaBP As Decimal = 0
        Dim diaAnterior As Date = Now.AddDays(-90)
        Dim FechaAplicacion As Date
        Dim cMinistracion As Integer
        Dim CFechaAutorizacion As String = ""
        Dim nPorcFega As Decimal = 0
        'diaAnterior = "12/02/2018"

        'llena fechas para detalle finagil
        cm4 = New SqlCommand("update mFINAGIL Set fechapago = fechaalta, fechadocumento = fechaalta where " _
            & "fechapago = '' and documento NOT IN ('EFECTIVO','REEMBOLSO') and fechaalta >= '" & diaAnterior.ToString("yyyyMMdd") & "'  ", cnAgil)
        cnAgil.Open()
        cm4.ExecuteScalar()
        cnAgil.Close()

        cm4 = New SqlCommand("SELECT Fecha FROM CONT_FechasAplicacion WHERE (Estatus = N'Vigente')", cnAgil)
        cnAgil.Open()
        FechaAplicacion = cm4.ExecuteScalar()
        cnAgil.Close()



        ' El siguiente Command trae todas las ministraciones que haya hecho FINAGIL en el mes de proceso
        ' "WHERE Avios.Ciclo IN ('05','06','07','08') AND FechaPago >= '20121208' AND FechaPago <= '20121212' AND Importe > 0 " & _
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT mFINAGIL.*, Cliente, Tipta, Tasas, DiferencialFINAGIL, Avios.fondeo, FechaAutorizacion, PorcFega, AplicaFega FROM mFINAGIL " &
                           "INNER JOIN Avios ON mFINAGIL.Anexo = Avios.Anexo AND mFINAGIL.Ciclo = Avios.Ciclo " &
                           "WHERE ((Avios.Ciclo >= '05' and Avios.tipar <> 'C') or (Avios.tipar = 'C')) AND " &
                           "FechaAlta >= '" & diaAnterior.ToString("yyyyMMdd") & "' AND (mFINAGIL.Notas = 'PAGADO')" &
                            "AND FechaPago <> '' And Importe > 0 AND (mFINAGIL.procesado is null or mFINAGIL.procesado <> 1) " &
                           "ORDER BY mFINAGIL.Ciclo, mFINAGIL.Anexo, FechaAlta, Ministracion"
            .Connection = cnAgil
        End With

        ' Llenar el DataSet lo cual abre y cierra la conexión

        daMinistracion.Fill(dsAgil, "Ministraciones")

        ' Tengo que crear una tabla donde voy a ir insertando las ministraciones.   Además, esta tabla debe contener
        ' una llave primaria para que pueda buscar un contrato en particular.

        For Each drMinistracion In dsAgil.Tables("Ministraciones").Rows
            Try

                cAnexo = drMinistracion("Anexo")
                cCiclo = drMinistracion("Ciclo")
                cCliente = drMinistracion("Cliente")
                cTipta = drMinistracion("Tipta")
                nTasa = drMinistracion("Tasas")
                cFondeo = drMinistracion("Fondeo")
                CFechaAutorizacion = drMinistracion("FechaAutorizacion")
                nPorcFega = drMinistracion("PorcFega")

                nDiferencial = drMinistracion("DiferencialFINAGIL")
                cFechaPago = drMinistracion("FechaPago")
                If cFechaPago < FechaAplicacion.ToString("yyyyMM01") Then
                    cFechaPago = FechaAplicacion.ToString("yyyyMM01")
                End If
                nImporte = drMinistracion("Importe")
                nGarantia = drMinistracion("Garantia")
                cDocumento = drMinistracion("Documento")
                cMinistracion = drMinistracion("Ministracion")

                With cm2
                    .CommandType = CommandType.Text
                    .CommandText = "SELECT * FROM DetalleFINAGIL " &
                                   "WHERE Anexo = '" & cAnexo & "' AND Ciclo = '" & cCiclo & "' " &
                                   "ORDER BY Consecutivo"
                    .Connection = cnAgil
                End With

                ' Llenar el DataSet lo cual abre y cierra la conexión

                daDetalle.Fill(dsAgil, "Detalle")

                If dsAgil.Tables("Detalle").Rows.Count = 0 Then

                    ' Es el primer registro de este contrato, al menos para el mes que se está procesando

                    nConsecutivo = 1
                    cFechaInicial = cFechaPago
                    cFechaFinal = cFechaPago
                    nDias = 0
                    nSaldoInicial = 0
                    nSaldoFinal = nImporte

                Else

                    ' Existen registros previos de este contrato por lo que tengo que tomar el dato más reciente
                    ' para determinar la Fecha Inicial y el Saldo Inicial

                    For Each drDetalle In dsAgil.Tables("Detalle").Rows
                        nConsecutivo = drDetalle("Consecutivo")
                        cFechaInicial = drDetalle("FechaFinal")
                        nSaldoInicial = drDetalle("SaldoFinal")
                    Next

                    nConsecutivo += 1
                    cFechaFinal = cFechaPago
                    nSaldoFinal = nSaldoInicial + nImporte

                    nDias = DateDiff(DateInterval.Day, CTOD(cFechaInicial), CTOD(cFechaFinal))

                End If

                If cTipta = "7" Then

                    nTasaBP = Round(nTasa + nDiferencial, 4)

                Else

                    ' Construyo una fecha que me permita buscar el promedio de la tasa TIIE del mes inmediato anterior

                    myKeySearch(0) = Mid(DTOC(DateAdd(DateInterval.Month, -1, CTOD(cFechaFinal))), 1, 6)

                    drTIIE = dtTIIE.Rows.Find(myKeySearch)

                    If drTIIE Is Nothing Then
                        nTasaBP = 0
                    Else
                        nTasaBP = drTIIE("Promedio")
                    End If

                    nTasaBP = Round(nTasaBP + nDiferencial, 4)

                End If
                If cFondeo = "03" Then
                    If CFechaAutorizacion >= "20160101" Then
                        If nPorcFega > 0 Then
                            nFEGA = Round(nImporte * nPorcFega, 2)
                        Else
                            nFEGA = Round(nImporte * 0.0174, 2)
                        End If

                    Else
                        nFEGA = Round(nImporte * 0.0116, 2)
                    End If
                    If drMinistracion("AplicaFega") = False Then
                        nFEGA = 0
                    End If
                Else
                    nFEGA = 0
                    nGarantia = 0
                End If

                nSaldoFinal = Round(nSaldoFinal + nFEGA + nGarantia, 2)

                strInsert = "INSERT INTO DetalleFINAGIL (Anexo, Ciclo, Cliente, Consecutivo, FechaInicial, FechaFinal, Dias, TasaBP, SaldoInicial, SaldoFinal, Concepto, Importe, FEGA, Garantia, Intereses,trdt,provinte) "
                strInsert = strInsert & "VALUES ('"
                strInsert = strInsert & cAnexo & "', '"
                strInsert = strInsert & cCiclo & "', '"
                strInsert = strInsert & cCliente & "', "
                strInsert = strInsert & nConsecutivo & ", '"
                strInsert = strInsert & cFechaInicial & "', '"
                strInsert = strInsert & cFechaFinal & "', "
                strInsert = strInsert & nDias & ", "
                strInsert = strInsert & nTasaBP & ", "
                strInsert = strInsert & nSaldoInicial & ", "
                strInsert = strInsert & nSaldoFinal & ", '"
                strInsert = strInsert & cDocumento & "', "
                strInsert = strInsert & nImporte & ", "
                strInsert = strInsert & nFEGA & ", "
                strInsert = strInsert & nGarantia & ", "
                strInsert = strInsert & nIntereses & ",'" & diaAnterior.ToString("MM/dd/yyyy") & "',1)"

                cm1 = New SqlCommand(strInsert, cnAgil)
                cm3 = New SqlCommand("update mFINAGIL Set Procesado = 1 where " _
                & "Anexo = '" & cAnexo & "' And Ciclo = '" & cCiclo & "'  " _
                & "and ministracion = " & cMinistracion & " and FechaPago = '" & cFechaPago & "';", cnAgil)
                cnAgil.Open()
                'MessageBox.Show(cm3.CommandText)
                ERRR.WriteLine(cm3.CommandText & "|" & Now.ToString)
                cm1.ExecuteNonQuery()
                cMinistracion = cm3.ExecuteScalar()
                cnAgil.Close()
                dsAgil.Tables.Remove("Detalle")
            Catch ex As Exception
                ERRR.WriteLine(ex.Message)
            End Try
        Next
        ERRR.WriteLine("proceso terminado " & Now.ToString)
        ERRR.Close()
        ERRR.Dispose()
    End Sub

    Public Function TIIEavg(ByVal cReferencia As String) As DataTable

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection("Server=SERVER-RAID; DataBase=Production; User ID = 'User_PRO'; pwd = 'User_PRO2015'")
        Dim cm1 As New SqlCommand()
        Dim daTIIE As New SqlDataAdapter(cm1)

        Dim dsAgil As New DataSet()
        Dim dtTIIEavg As New DataTable()
        Dim drTasa As DataRow
        Dim drTemporal As DataRow
        Dim myColArray(1) As DataColumn
        Dim myKeySearch(0) As String

        ' Declaración de variables de datos

        Dim cMes As String
        Dim nValor As Decimal = 0

        If cReferencia = "FINAGIL" Then

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT SUBSTRING(Vigencia,1,6) AS Mes, ROUND(AVG(Valor),4) AS Promedio FROM Hista " &
                               "WHERE Tasa = '4' " &
                               "GROUP BY SUBSTRING(Vigencia,1,6) " &
                               "ORDER BY SUBSTRING(Vigencia,1,6)"
                .Connection = cnAgil
            End With

            ' Llenar el dataset lo cual abre y cierra la conexión

            daTIIE.Fill(dsAgil, "TIIE")

            ' Tengo que definir una llave primaria para la tabla

            myColArray(0) = dsAgil.Tables("TIIE").Columns("Mes")
            dsAgil.Tables("TIIE").PrimaryKey = myColArray

        ElseIf cReferencia = "FIRA" Then

            ' Primero creo la tabla dtTIIEavg

            dtTIIEavg.Columns.Add("Mes", Type.GetType("System.String"))
            dtTIIEavg.Columns.Add("Promedio", Type.GetType("System.Decimal"))
            dtTIIEavg.Columns.Add("Suma", Type.GetType("System.Decimal"))
            dtTIIEavg.Columns.Add("DiasHabiles", Type.GetType("System.Decimal"))

            ' Tengo que definir una llave primaria para la tabla dtTIIEavg a fin de buscar un anexo
            ' para acumular ministraciones

            myColArray(0) = dtTIIEavg.Columns("Mes")
            dtTIIEavg.PrimaryKey = myColArray

            '  Para el promedio NO tengo que considerar la TIIE de sábados ni domingos, ni de días festivos oficiales

            With cm1
                .CommandType = CommandType.Text
                .CommandText = "SELECT * FROM Hista " &
                               "WHERE Tasa = '4' " &
                               "ORDER BY Vigencia"
                .Connection = cnAgil
            End With

            ' Llenar el dataset lo cual abre y cierra la conexión

            daTIIE.Fill(dsAgil, "TIIE")

            For Each drTasa In dsAgil.Tables("TIIE").Rows
                If drTasa("Festivo") <> "S" And Weekday(CTOD(drTasa("Vigencia"))) <> 1 And Weekday(CTOD(drTasa("Vigencia"))) <> 7 Then
                    cMes = Mid(drTasa("Vigencia"), 1, 6)
                    nValor = drTasa("Valor")
                    myKeySearch(0) = cMes
                    drTemporal = dtTIIEavg.Rows.Find(myKeySearch)
                    If drTemporal Is Nothing Then
                        drTemporal = dtTIIEavg.NewRow()
                        drTemporal("Mes") = cMes
                        drTemporal("Promedio") = 0
                        drTemporal("Suma") = nValor
                        drTemporal("DiasHabiles") = 1
                        dtTIIEavg.Rows.Add(drTemporal)
                    Else
                        drTemporal("Suma") += nValor
                        drTemporal("DiasHabiles") += 1
                    End If
                End If
            Next

            For Each drTasa In dtTIIEavg.Rows
                drTasa("Promedio") = Round(drTasa("Suma") / drTasa("DiasHabiles"), 4)
            Next

            dsAgil.Tables.Remove("TIIE")
            dsAgil.Tables.Add(dtTIIEavg)

        End If

        TIIEavg = dsAgil.Tables(0)

        cnAgil.Dispose()
        cm1.Dispose()

    End Function

    Private Sub btnCierre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCierre.Click

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection("Server=SERVER-RAID; DataBase=Production; User ID = 'User_PRO'; pwd = 'User_PRO2015'")
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim daCreditos As New SqlDataAdapter(cm1)
        Dim daDetalle As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet()
        Dim drCredito As DataRow
        Dim drDetalle As DataRow

        Dim strInsert As String
        Dim strUpdate As String

        ' Declaración de variables de datos

        Dim cAnexo As String = ""
        Dim cCiclo As String = ""
        Dim cCliente As String = ""
        Dim cDocumento As String = "INTERESES"
        Dim cFecha As String = "20140825"
        Dim cFechaFinal As String = ""
        Dim cFechaInicial As String = ""
        Dim cFechaPago As String = ""
        Dim cTipta As String = ""
        Dim cUltimoCorte As String = ""
        Dim nConsecutivo As Integer = 0
        Dim nDias As Integer = 0
        Dim nDiferencial As Decimal = 0
        Dim nIntereses As Decimal = 0
        Dim nSaldoFEGA As Decimal = 0
        Dim nSaldoFinal As Decimal = 0
        Dim nSaldoGarantia As Decimal = 0
        Dim nSaldoImporte As Decimal = 0
        Dim nSaldoIntereses As Decimal = 0
        Dim nSaldoInicial As Decimal = 0
        Dim nSumaIntereses As Decimal = 0
        Dim nTasa As Decimal = 0
        Dim nTasaBP As Decimal = 0

        ' El siguiente Command trae todos los contratos que tengan saldo

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo FROM DetalleFINAGIL " &
                           "INNER JOIN Avios ON DetalleFINAGIL.Anexo = Avios.Anexo AND DetalleFINAGIL.Ciclo = Avios.Ciclo " &
                           "INNER JOIN Clientes ON Avios.Cliente = Clientes.Cliente " &
                           "WHERE Avios.Ciclo >= '05' " &
                           "GROUP BY DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo " &
                           "HAVING SUM(Importe+FEGA+Garantia) > 0 " &
                           "ORDER BY DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo"
            .Connection = cnAgil
        End With

        ' Llenar el DataSet lo cual abre y cierra la conexión

        daCreditos.Fill(dsAgil, "Creditos")


        ' Tengo que crear una tabla donde voy a ir insertando las ministraciones.   Además, esta tabla debe contener
        ' una llave primaria para que pueda buscar un contrato en particular.

        For Each drCredito In dsAgil.Tables("Creditos").Rows

            cAnexo = drCredito("Anexo")
            cCiclo = drCredito("Ciclo")

            ' El siguiente Command trae los movimientos que existan en DetalleFINAGIL del contrato seleccionado

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "SELECT DetalleFINAGIL.*, Tipta, Tasas, DiferencialFINAGIL, UltimoCorte FROM DetalleFINAGIL " &
                               "INNER JOIN Avios ON DetalleFINAGIL.Anexo = Avios.Anexo AND DetalleFINAGIL.Ciclo = Avios.Ciclo " &
                               "WHERE DetalleFINAGIL.Anexo = '" & cAnexo & "' AND DetalleFINAGIL.Ciclo = '" & cCiclo & "' " &
                               "ORDER BY DetalleFINAGIL.Anexo, Consecutivo"
                .Connection = cnAgil
            End With

            ' Llenar el DataSet lo cual abre y cierra la conexión

            daDetalle.Fill(dsAgil, "Detalle")

            For Each drDetalle In dsAgil.Tables("Detalle").Rows

                nTasa = drDetalle("Tasas")
                nDiferencial = drDetalle("DiferencialFINAGIL")
                cTipta = drDetalle("Tipta")
                cUltimoCorte = drDetalle("UltimoCorte")

                cCliente = drDetalle("Cliente")
                nConsecutivo = drDetalle("Consecutivo")
                cFechaInicial = drDetalle("FechaFinal")
                nSaldoInicial = drDetalle("SaldoFinal")

            Next

            nConsecutivo += 1

            If cTipta = "7" Then

                nTasaBP = Round(nTasa + nDiferencial, 4)

            Else

                ' Construyo una fecha que me permita buscar el promedio de la tasa TIIE del mes inmediato anterior

                myKeySearch(0) = Mid(DTOC(DateAdd(DateInterval.Month, -1, CTOD(cFecha))), 1, 6)

                drTIIE = dtTIIE.Rows.Find(myKeySearch)

                If drTIIE Is Nothing Then
                    nTasaBP = 0
                Else
                    nTasaBP = drTIIE("Promedio")
                End If

                nTasaBP = Round(nTasaBP + nDiferencial, 4)

            End If

            nDias = DateDiff(DateInterval.Day, CTOD(cFechaInicial), CTOD(cFecha))

            drDetalle = dsAgil.Tables("Detalle").NewRow
            drDetalle("Anexo") = cAnexo
            drDetalle("Anexo") = cCiclo
            drDetalle("Cliente") = cCliente
            drDetalle("Consecutivo") = nConsecutivo
            drDetalle("FechaInicial") = cFechaInicial
            drDetalle("FechaFinal") = cFecha
            drDetalle("Dias") = nDias
            drDetalle("TasaBP") = nTasaBP
            drDetalle("SaldoInicial") = nSaldoInicial
            drDetalle("SaldoFinal") = 0
            drDetalle("Concepto") = "INTERESES"
            drDetalle("Importe") = 0
            drDetalle("FEGA") = 0
            drDetalle("Garantia") = 0
            drDetalle("Intereses") = 0
            dsAgil.Tables("Detalle").Rows.Add(drDetalle)

            nSumaIntereses = 0

            For Each drDetalle In dsAgil.Tables("Detalle").Rows

                cFechaFinal = drDetalle("FechaFinal")
                If Mid(cFechaFinal, 1, 6) = Mid(cFecha, 1, 6) And cFechaFinal > cUltimoCorte Then
                    nSaldoInicial = drDetalle("SaldoInicial")
                    nTasaBP = drDetalle("TasaBP")
                    nDias = drDetalle("Dias")
                    nIntereses = Round(nSaldoInicial * nTasaBP / 36000 * nDias, 2)
                    nSumaIntereses = Round(nSumaIntereses + nIntereses, 2)
                End If

                nConsecutivo = drDetalle("Consecutivo")

            Next

            nSaldoFinal = nSaldoInicial + nSumaIntereses

            drDetalle("Intereses") = nSumaIntereses
            drDetalle("SaldoFinal") = nSaldoFinal

            If drDetalle("SaldoInicial") = 0 And drDetalle("SaldoFinal") = 0 Then
                dsAgil.Tables("Detalle").Rows(nConsecutivo - 1).Delete()
            ElseIf drDetalle("Importe") = 0 And drDetalle("FEGA") = 0 And drDetalle("Garantia") = 0 And drDetalle("Intereses") = 0 Then
                dsAgil.Tables("Detalle").Rows(nConsecutivo - 1).Delete()
            Else

                cnAgil.Open()

                strInsert = "INSERT INTO DetalleFINAGIL (Anexo, Ciclo, Cliente, Consecutivo, FechaInicial, FechaFinal, Dias, TasaBP, SaldoInicial, SaldoFinal, Concepto, Importe, FEGA, Garantia, Intereses) "
                strInsert = strInsert & "VALUES ('"
                strInsert = strInsert & cAnexo & "', '"
                strInsert = strInsert & cCiclo & "', '"
                strInsert = strInsert & cCliente & "', "
                strInsert = strInsert & nConsecutivo & ", '"
                strInsert = strInsert & cFechaInicial & "', '"
                strInsert = strInsert & cFechaFinal & "', "
                strInsert = strInsert & nDias & ", "
                strInsert = strInsert & nTasaBP & ", "
                strInsert = strInsert & nSaldoInicial & ", "
                strInsert = strInsert & nSaldoFinal & ", '"
                strInsert = strInsert & cDocumento & "', "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & nSumaIntereses & ")"
                cm1 = New SqlCommand(strInsert, cnAgil)
                cm1.ExecuteNonQuery()

                strUpdate = "UPDATE Avios SET UltimoCorte = '" & cFechaFinal & "' WHERE Anexo = '" & cAnexo & "' AND Ciclo = '" & cCiclo & "'"
                cm1 = New SqlCommand(strUpdate, cnAgil)
                cm1.ExecuteNonQuery()

                cnAgil.Close()

            End If

            dsAgil.Tables.Remove("Detalle")

        Next

        MsgBox("Cierre terminado", MsgBoxStyle.Information, "Mensaje del Sistema")

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection("Server=SERVER-RAID; DataBase=Production; User ID = 'User_PRO'; pwd = 'User_PRO2015'")
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim daCreditos As New SqlDataAdapter(cm1)
        Dim daDetalle As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet()
        Dim drCredito As DataRow
        Dim drDetalle As DataRow

        Dim strInsert As String
        Dim strUpdate As String

        ' Declaración de variables de datos

        Dim cAnexo As String = ""
        Dim cCiclo As String = ""
        Dim cCliente As String = ""
        Dim cDocumento As String = "INTERESES"
        Dim diaAnterior As Date = Now.AddDays(-1)
        Dim cFecha As String = diaAnterior.ToString("yyyyMMdd")
        Dim cFechaFinal As String = ""
        Dim cFechaInicial As String = ""
        Dim cFechaPago As String = ""
        Dim cTipta As String = ""
        Dim cUltimoCorte As String = ""
        Dim nConsecutivo As Integer = 0
        Dim nDias As Integer = 0
        Dim nDiferencial As Decimal = 0
        Dim nIntereses As Decimal = 0
        Dim nSaldoFEGA As Decimal = 0
        Dim nSaldoFinal As Decimal = 0
        Dim nSaldoGarantia As Decimal = 0
        Dim nSaldoImporte As Decimal = 0
        Dim nSaldoIntereses As Decimal = 0
        Dim nSaldoInicial As Decimal = 0
        Dim nSumaIntereses As Decimal = 0
        Dim nTasa As Decimal = 0
        Dim nTasaBP As Decimal = 0

        cFecha = "20140807"

        ' El siguiente Command trae todos los contratos que tengan saldo

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo, fechaterminacion FROM DetalleFINAGIL " &
                            "INNER JOIN Avios ON DetalleFINAGIL.Anexo = Avios.Anexo AND DetalleFINAGIL.Ciclo = Avios.Ciclo " &
                            "INNER JOIN Clientes ON Avios.Cliente = Clientes.Cliente " &
                            "WHERE (Avios.Ciclo >= '05' and Avios.tipar = 'H') or (Avios.tipar = 'C') " &
                            "GROUP BY DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo, fechaterminacion " &
                            "HAVING SUM(Importe+FEGA+Garantia) > 0 and fechaterminacion = '" & cFecha & "' " &
                            "ORDER BY DetalleFINAGIL.Anexo, DetalleFINAGIL.Ciclo "
            .Connection = cnAgil
        End With

        ' Llenar el DataSet lo cual abre y cierra la conexión

        daCreditos.Fill(dsAgil, "Creditos")


        ' Tengo que crear una tabla donde voy a ir insertando las ministraciones.   Además, esta tabla debe contener
        ' una llave primaria para que pueda buscar un contrato en particular.

        For Each drCredito In dsAgil.Tables("Creditos").Rows

            cAnexo = drCredito("Anexo")
            cCiclo = drCredito("Ciclo")

            ' El siguiente Command trae los movimientos que existan en DetalleFINAGIL del contrato seleccionado

            With cm2
                .CommandType = CommandType.Text
                .CommandText = "SELECT DetalleFINAGIL.*, Tipta, Tasas, DiferencialFINAGIL, UltimoCorte FROM DetalleFINAGIL " &
                               "INNER JOIN Avios ON DetalleFINAGIL.Anexo = Avios.Anexo AND DetalleFINAGIL.Ciclo = Avios.Ciclo " &
                               "WHERE DetalleFINAGIL.Anexo = '" & cAnexo & "' AND DetalleFINAGIL.Ciclo = '" & cCiclo & "' " &
                               "ORDER BY DetalleFINAGIL.Anexo, Consecutivo"
                .Connection = cnAgil
            End With

            ' Llenar el DataSet lo cual abre y cierra la conexión

            daDetalle.Fill(dsAgil, "Detalle")

            For Each drDetalle In dsAgil.Tables("Detalle").Rows

                nTasa = drDetalle("Tasas")
                nDiferencial = drDetalle("DiferencialFINAGIL")
                cTipta = drDetalle("Tipta")
                cUltimoCorte = drDetalle("UltimoCorte")

                cCliente = drDetalle("Cliente")
                nConsecutivo = drDetalle("Consecutivo")
                cFechaInicial = drDetalle("FechaFinal")
                nSaldoInicial = drDetalle("SaldoFinal")

            Next

            nConsecutivo += 1

            If cTipta = "7" Then

                nTasaBP = Round(nTasa + nDiferencial, 4)

            Else

                ' Construyo una fecha que me permita buscar el promedio de la tasa TIIE del mes inmediato anterior

                myKeySearch(0) = Mid(DTOC(DateAdd(DateInterval.Month, -1, CTOD(cFecha))), 1, 6)

                drTIIE = dtTIIE.Rows.Find(myKeySearch)

                If drTIIE Is Nothing Then
                    nTasaBP = 0
                Else
                    nTasaBP = drTIIE("Promedio")
                End If

                nTasaBP = Round(nTasaBP + nDiferencial, 4)

            End If

            nDias = DateDiff(DateInterval.Day, CTOD(cFechaInicial), CTOD(cFecha))

            drDetalle = dsAgil.Tables("Detalle").NewRow
            drDetalle("Anexo") = cAnexo
            drDetalle("Anexo") = cCiclo
            drDetalle("Cliente") = cCliente
            drDetalle("Consecutivo") = nConsecutivo
            drDetalle("FechaInicial") = cFechaInicial
            drDetalle("FechaFinal") = cFecha
            drDetalle("Dias") = nDias
            drDetalle("TasaBP") = nTasaBP
            drDetalle("SaldoInicial") = nSaldoInicial
            drDetalle("SaldoFinal") = 0
            drDetalle("Concepto") = "INTERESES"
            drDetalle("Importe") = 0
            drDetalle("FEGA") = 0
            drDetalle("Garantia") = 0
            drDetalle("Intereses") = 0
            dsAgil.Tables("Detalle").Rows.Add(drDetalle)

            nSumaIntereses = 0

            For Each drDetalle In dsAgil.Tables("Detalle").Rows

                cFechaFinal = drDetalle("FechaFinal")
                If Mid(cFechaFinal, 1, 6) = Mid(cFecha, 1, 6) And cFechaFinal > cUltimoCorte Then
                    nSaldoInicial = drDetalle("SaldoInicial")
                    nTasaBP = drDetalle("TasaBP")
                    nDias = drDetalle("Dias")
                    nIntereses = Round(nSaldoInicial * nTasaBP / 36000 * nDias, 2)
                    nSumaIntereses = Round(nSumaIntereses + nIntereses, 2)
                End If

                nConsecutivo = drDetalle("Consecutivo")

            Next

            nSaldoFinal = nSaldoInicial + nSumaIntereses

            drDetalle("Intereses") = nSumaIntereses
            drDetalle("SaldoFinal") = nSaldoFinal

            If drDetalle("SaldoInicial") = 0 And drDetalle("SaldoFinal") = 0 Then
                dsAgil.Tables("Detalle").Rows(nConsecutivo - 1).Delete()
            ElseIf drDetalle("Importe") = 0 And drDetalle("FEGA") = 0 And drDetalle("Garantia") = 0 And drDetalle("Intereses") = 0 Then
                dsAgil.Tables("Detalle").Rows(nConsecutivo - 1).Delete()
            Else

                cnAgil.Open()

                strInsert = "INSERT INTO DetalleFINAGIL (Anexo, Ciclo, Cliente, Consecutivo, FechaInicial, FechaFinal, Dias, TasaBP, SaldoInicial, SaldoFinal, Concepto, Importe, FEGA, Garantia, Intereses) "
                strInsert = strInsert & "VALUES ('"
                strInsert = strInsert & cAnexo & "', '"
                strInsert = strInsert & cCiclo & "', '"
                strInsert = strInsert & cCliente & "', "
                strInsert = strInsert & nConsecutivo & ", '"
                strInsert = strInsert & cFechaInicial & "', '"
                strInsert = strInsert & cFechaFinal & "', "
                strInsert = strInsert & nDias & ", "
                strInsert = strInsert & nTasaBP & ", "
                strInsert = strInsert & nSaldoInicial & ", "
                strInsert = strInsert & nSaldoFinal & ", '"
                strInsert = strInsert & cDocumento & "', "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & 0 & ", "
                strInsert = strInsert & nSumaIntereses & ")"
                cm1 = New SqlCommand(strInsert, cnAgil)
                cm1.ExecuteNonQuery()

                strUpdate = "UPDATE Avios SET UltimoCorte = '" & cFechaFinal & "' WHERE Anexo = '" & cAnexo & "' AND Ciclo = '" & cCiclo & "'"
                cm1 = New SqlCommand(strUpdate, cnAgil)
                cm1.ExecuteNonQuery()

                cnAgil.Close()

            End If

            dsAgil.Tables.Remove("Detalle")

        Next

        '       MsgBox("Cierre terminado", MsgBoxStyle.Information, "Mensaje del Sistema")
        '
        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()

    End Sub
End Class
