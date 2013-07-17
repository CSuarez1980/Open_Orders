Imports SAPCOM.RepairsLevels
Imports OAConnection

Public Class Form1
    Public Status As String
    Private WithEvents X As New Manager

    Public Event M(ByVal MSG)
    Public D As New DataTable
    Public res As Integer = 10800



    Public Sub SM(ByVal msg As String) Handles Me.M
        'D.Columns.Add("Message")

        Dim r As DataRow = D.NewRow
        r("Message") = msg

        D.Rows.Add(r)
        ' DG.DataSource = D


        'Dim ms As String = ""

        'ms = msg

        'Me.lstStatus.Items.Add(ms)


        'MsgBox(msg)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EndProcess()

        Dim cn As New OAConnection.Connection
        pgrWorking.Style = ProgressBarStyle.Marquee
        lstStatus.Items.Add(Now.ToString & "-" & "Cleaning Open Orders Table")


        GetReports()


        'cn.ExecuteInServer("Delete From SC_OpenOrders")
        'cn.ExecuteInServer("Delete From DMS_Confirmation")

        'AddHandler X.MyReport, AddressOf SM
        'D.Columns.Add("Message")

        ' DG.DataSource




        'BS.DataSource = D
        'BS.ResumeBinding()
        'DG.DataBindings.Add("DataSource", X, "D", True, DataSourceUpdateMode.OnPropertyChanged)
        'lstStatus.DataBindings.Add("DataSource", X, "D", True, DataSourceUpdateMode.OnPropertyChanged)


        'BGW.RunWorkerAsync()

        ''BGOI_L7P.RunWorkerAsync()
        ''BGOI_G4P.RunWorkerAsync()
        ''BGOI_GBP.RunWorkerAsync()
        ''BGOI_L6P.RunWorkerAsync()
    End Sub

    Private Sub EndProcess()
        Dim cn As New OAConnection.Connection
        cn.ExecuteInServer("DELETE FROM SC_OpenOrders WHERE ([Doc Type] <> 'EC') AND (SAPBox <> 'L7P') AND (NOT EXISTS (SELECT Name, Tnumber, Status FROM dbo.[PSS People] WHERE (dbo.SC_OpenOrders.[Created By] = Tnumber)))")
        cn.ExecuteInServer("DELETE FROM SC_OpenOrders WHERE ([Doc Type] = 'NB') AND (SAPBox = 'L7P') AND (NOT EXISTS (SELECT Name, Tnumber, Status FROM dbo.[PSS People] WHERE (dbo.SC_OpenOrders.[Created By] = Tnumber)))")
        cn.ExecuteInServer("DELETE From SC_OpenOrders Where (Vendor = '15145463')")

        'To Do:
        'Eliminar las PO de la distribucion temporal que ya no estan abiertas
        cn.ExecuteInServer("DELETE FROM dbo.LA_TMP_Open_Orders_Distribution Where (NOT EXISTS (SELECT [Doc Number] FROM SC_OpenOrders Where (dbo.LA_TMP_Open_Orders_Distribution.[Doc Number] = [Doc Number]) AND (dbo.LA_TMP_Open_Orders_Distribution.SAPBox = SAPBox)))")

        'Agregar las POs que son nuevas a la distribucion temporal
        cn.ExecuteInServer("Insert Into LA_TMP_Open_Orders_Distribution (SAPBox, [Doc Number]) SELECT DISTINCT SAPBox, [Doc Number] From SC_OpenOrders WHERE (NOT EXISTS (SELECT [Doc Number] From dbo.LA_TMP_Open_Orders_Distribution Where (dbo.SC_OpenOrders.[Doc Number] = [Doc Number]) AND (dbo.SC_OpenOrders.SAPBox = SAPBox)))")

        'Crear una funcion para asignarles el owner a las nuevas.
        Dim OO As New DataTable
        OO = cn.RunSentence("Select * From vst_LA_Check_Distribution").Tables(0)

        If OO.Rows.Count > 0 Then
            For Each r As DataRow In OO.Rows
                Try


                    Dim OI As New OAConnection.DMS_User(r("SAPBox"), r("Mat Group"), r("Purch Grp"), r("Purch Org"), r("Plant"))
                    OI.Execute()

                    If OI.Success Then
                        cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = '" & OI.SPS & "', Owner = '" & OI.PTB & "' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    Else
                        cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = 'BB0898', Owner = 'BB0898' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    End If


                    'Dim lO As New Owner
                    'lO = GetOwner(r("SAPBox"), r("Spend"), r("Plant"), r("Purch Grp"), r("Purch Org"))
                    'If Not lO Is Nothing Then
                    '    cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = '" & lO.SPS & "', Owner = '" & lO.Owner & "' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    'Else
                    '    'Si no tiene Owner se le asigna a Alejandra Baltodano
                    '    cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = 'BB0898', Owner = 'BB0898' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    'End If
                Catch ex As Exception

                End Try
            Next
        End If

        'Actualización de los vendors en la tabla VendorsG11.
        Dim Vn As New SAPCOM.LFA1_Report("G4P", "BM4691", "LAT")
        Dim NV As New DataTable 'New vendors

        NV = cn.RunSentence("Select * From vst_New_Vendors").Tables(0)
        If NV.Rows.Count > 0 Then
            For Each v In NV.Rows
                Vn.IncludeVendor(v("Vendor"))
            Next

            Vn.Execute()
            If Vn.Success Then

                For Each v In NV.Rows
                    Try
                        Dim VR = (From C In Vn.Data.AsEnumerable() _
                                  Where C.Item("Vendor") = v("Vendor") _
                                  Select C.Item("Country")).First

                        v("Country") = VR
                    Catch ex As Exception
                        'Do nothing
                    End Try
                Next

                cn.AppendTableToSqlServer("VendorsG11", NV)
            End If
        End If

        End
    End Sub
    'Private Sub BGOI_L7P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOI_L7P.DoWork
    '    Dim SAP As String = "L7P"
    '    Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)
    '    Dim POs As New DataTable
    '    Dim cn As New OAConnection.Connection
    '    Dim GetConfirmation As Boolean
    '    dtPlants.Tabla = "SC_Plant"
    '    dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
    '    dtPlants.Execute()

    '    For Each row As DataRow In dtPlants.Data.Rows
    '        GetConfirmation = False
    '        Dim Rep As New SAPCOM.OpenOrders_Report(SAP, "BM4691", "LAT")
    '        Status = "Getting: " & SAP & " - Plant: " & row("Plant_Code")
    '        BGOI_L7P.ReportProgress(0)

    '        Rep.RepairsLevel = IncludeRepairs
    '        Rep.Include_GR_IR = True
    '        Rep.IncludeDelivDates = True
    '        Rep.IncludePlant(row("Plant_Code"))

    '        Rep.Execute()
    '        Dim EKES As New SAPCOM.EKES_Report(SAP, "BM4691", "LAT")
    '        Dim EKKO As New SAPCOM.EKKO_Report(SAP, "BM4691", "LAT")

    '        If Rep.Success Then
    '            If Rep.ErrMessage = Nothing Then
    '                POs = Rep.Data
    '                GetConfirmation = True
    '                If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0219")
    '                End If
    '                If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '                    POs.Columns.Remove("EKPO-ZWERT")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0218")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0220")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '                    POs.Columns.Remove("EKKO-MEMORYTYPE")
    '                End If

    '                Dim TN As New DataColumn
    '                Dim SB As New DataColumn

    '                TN.ColumnName = "Usuario"
    '                TN.Caption = "Usuario"
    '                TN.DefaultValue = "BM4691"

    '                SB.DefaultValue = SAP
    '                SB.ColumnName = "SAPBox"
    '                SB.Caption = "SAPBox"

    '                POs.Columns.Add(TN)
    '                POs.Columns.Add(SB)

    '                cn.AppendTableToSqlServer("SC_OpenOrders", POs)

    '                Dim cRow As DataRow
    '                For Each cRow In POs.Rows
    '                    EKKO.IncludeDocument(cRow.Item("Doc Number"))
    '                Next

    '                For Each cRow In POs.Rows
    '                    EKES.IncludeDocument(cRow.Item("Doc Number"))
    '                Next

    '            Else
    '                Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '                BGOI_L7P.ReportProgress(0)
    '            End If
    '        Else
    '            Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '            BGOI_L7P.ReportProgress(0)
    '        End If



    '        '****************************************************************
    '        'Las POs GR103 ya no deben mostarse en el reporte al formar parte de ITrade


    '        'Dim Rep2 As New SAPCOM.OpenGR105_Report(SAP, "BM4691", "LAT")
    '        'Status = "Getting: " & SAP & " - GR 103 Plant: " & row("Plant_Code")

    '        'BGOI_L7P.ReportProgress(0)

    '        'Rep2.RepairsLevel = IncludeRepairs
    '        'Rep2.Include_GR_IR = True
    '        'Rep2.IncludeDelivDates = True
    '        'Rep2.IncludePlant(row("Plant_Code"))

    '        'Rep2.Execute()
    '        'If Rep2.Success Then
    '        '    If Rep2.ErrMessage = Nothing Then
    '        '        POs = Rep2.Data
    '        '        GetConfirmation = True

    '        '        If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '        '            POs.Columns.Remove("EKKO-WAERS-0219")
    '        '        End If
    '        '        If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '        '            POs.Columns.Remove("EKPO-ZWERT")
    '        '        End If
    '        '        If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '        '            POs.Columns.Remove("EKKO-WAERS-0218")
    '        '        End If
    '        '        If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '        '            POs.Columns.Remove("EKKO-WAERS-0220")
    '        '        End If
    '        '        If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '        '            POs.Columns.Remove("EKKO-MEMORYTYPE")
    '        '        End If

    '        '        Dim cRow As DataRow
    '        '        For Each cRow In Rep2.Data.Rows
    '        '            EKKO.IncludeDocument(cRow.Item("Doc Number"))
    '        '        Next

    '        '        For Each cRow In Rep2.Data.Rows
    '        '            EKES.IncludeDocument(cRow.Item("Doc Number"))
    '        '        Next

    '        '        Dim TN2 As New DataColumn
    '        '        Dim SB2 As New DataColumn

    '        '        TN2.ColumnName = "Usuario"
    '        '        TN2.Caption = "Usuario"
    '        '        TN2.DefaultValue = "BM4691"

    '        '        SB2.DefaultValue = SAP
    '        '        SB2.ColumnName = "SAPBox"
    '        '        SB2.Caption = "SAPBox"

    '        '        POs.Columns.Add(TN2)
    '        '        POs.Columns.Add(SB2)

    '        '        Try
    '        '            cn.AppendTableToSqlServer("SC_OpenOrders", POs)
    '        '        Catch ex As Exception
    '        '            InsertRowByRow(POs)
    '        '        End Try
    '        '    End If
    '        'End If

    '        If GetConfirmation Then
    '            Status = "Getting: " & SAP & " - Confirmation: " & row("Plant_Code")
    '            BGOI_L7P.ReportProgress(0)

    '            Status = "Getting: " & SAP & " - EKES: " & row("Plant_Code")
    '            BGOI_L7P.ReportProgress(0)

    '            EKES.Execute()
    '            If EKES.Success Then
    '                Dim SBE As New DataColumn
    '                SBE.DefaultValue = SAP
    '                SBE.ColumnName = "SAPBox"
    '                SBE.Caption = "SAPBox"
    '                EKES.Data.Columns.Add(SBE)
    '                cn.AppendTableToSqlServer("DMS_Confirmation", EKES.Data)
    '                'cn.ExecuteInServer("Delete From DMS_Confirmation Where Confirmation = 'YV'")
    '            Else
    '                Status = "-> Fail: " & SAP & " - EKES: " & row("Plant_Code") & "[" & EKES.ErrMessage & "]"
    '                BGOI_L7P.ReportProgress(0)
    '            End If

    '            Status = "Getting: " & SAP & " - EKKO: " & row("Plant_Code")
    '            BGOI_L7P.ReportProgress(0)

    '            EKKO.Execute()

    '            If EKKO.Success Then
    '                Dim ESB As New DataColumn
    '                ESB.DefaultValue = SAP
    '                ESB.ColumnName = "SAPBox"
    '                ESB.Caption = "SAPBox"

    '                If EKKO.Data.Columns.IndexOf("Company Code") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Company Code")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Doc Type") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Doc Type")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Created On") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Created On")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Created By") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Created By")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Vendor") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Vendor")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Language") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Language")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("POrg") <> -1 Then
    '                    EKKO.Data.Columns.Remove("POrg")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("PGrp") <> -1 Then
    '                    EKKO.Data.Columns.Remove("PGrp")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Currency") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Currency")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Doc Date") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Doc Date")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Validity Start") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Validity Start")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Validity End") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Validity End")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Y Refer") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Y Refer")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("SalesPerson") <> -1 Then
    '                    EKKO.Data.Columns.Remove("SalesPerson")
    '                End If
    '                If EKKO.Data.Columns.IndexOf("Telephone") <> -1 Then
    '                    EKKO.Data.Columns.Remove("Telephone")
    '                End If

    '                EKKO.Data.Columns.Add(ESB)

    '                For Each r As DataRow In EKKO.Data.Rows
    '                    If r("O Reference").ToString.IndexOf("Y") <> -1 Then
    '                        r("O Reference") = "Manual"
    '                    Else
    '                        r("O Reference") = ""
    '                    End If
    '                Next

    '                cn.AppendTableToSqlServer("DMS_Confirmation", EKKO.Data)
    '                cn.ExecuteInServer("Delete From DMS_Confirmation Where Confirmation = ''")
    '            End If
    '        End If
    '    Next
    'End Sub
    'Private Sub BGOI_L7P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOI_L7P.ProgressChanged
    '    lblStatus.Text = Status
    '    lstStatus.Items.Add(Status)
    'End Sub
    'Private Sub BGOI_L7P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOI_L7P.RunWorkerCompleted
    '    EndProcess()
    'End Sub
    'Private Sub BGOI_GBP_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOI_GBP.DoWork
    '    Dim SAP As String = "GBP"
    '    Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)

    '    Dim POs As New DataTable
    '    Dim cn As New OAConnection.Connection
    '    dtPlants.Tabla = "SC_Plant"
    '    dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
    '    dtPlants.Execute()

    '    For Each row As DataRow In dtPlants.Data.Rows
    '        Dim Rep As New SAPCOM.OpenOrders_Report(SAP, "BM4691", "LAT")
    '        Status = "Getting: " & SAP & " - Plant: " & row("Plant_Code")

    '        BGOI_GBP.ReportProgress(0)

    '        Rep.RepairsLevel = IncludeRepairs
    '        Rep.Include_GR_IR = True
    '        Rep.IncludeDelivDates = True
    '        Rep.IncludePlant(row("Plant_Code"))

    '        Rep.Execute()
    '        If Rep.Success Then
    '            If Rep.ErrMessage = Nothing Then
    '                POs = Rep.Data
    '                If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0219")
    '                End If
    '                If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '                    POs.Columns.Remove("EKPO-ZWERT")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0218")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0220")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '                    POs.Columns.Remove("EKKO-MEMORYTYPE")
    '                End If

    '                Dim TN As New DataColumn
    '                Dim SB As New DataColumn

    '                TN.ColumnName = "Usuario"
    '                TN.Caption = "Usuario"
    '                TN.DefaultValue = "BM4691"

    '                SB.DefaultValue = SAP
    '                SB.ColumnName = "SAPBox"
    '                SB.Caption = "SAPBox"

    '                POs.Columns.Add(TN)
    '                POs.Columns.Add(SB)

    '                cn.AppendTableToSqlServer("SC_OpenOrders", POs)
    '            Else
    '                Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '                BGOI_GBP.ReportProgress(0)
    '            End If
    '        Else
    '            Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '            BGOI_GBP.ReportProgress(0)
    '        End If
    '    Next
    'End Sub
    'Private Sub BGOI_GBP_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOI_GBP.ProgressChanged
    '    lblStatus.Text = Status
    '    lstStatus.Items.Add(Status)
    'End Sub
    'Private Sub BGOI_GBP_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOI_GBP.RunWorkerCompleted
    '    EndProcess()
    'End Sub
    'Private Sub BGOI_L6P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOI_L6P.DoWork
    '    Dim SAP As String = "L6P"
    '    Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)

    '    Dim POs As New DataTable
    '    Dim cn As New OAConnection.Connection
    '    dtPlants.Tabla = "SC_Plant"
    '    dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
    '    dtPlants.Execute()

    '    For Each row As DataRow In dtPlants.Data.Rows
    '        Dim Rep As New SAPCOM.OpenOrders_Report(SAP, "BM4691", "LAT")
    '        Status = "Getting: " & SAP & " - Plant: " & row("Plant_Code")

    '        BGOI_L6P.ReportProgress(0)

    '        Rep.RepairsLevel = IncludeRepairs
    '        Rep.Include_GR_IR = True
    '        Rep.IncludeDelivDates = True
    '        Rep.IncludePlant(row("Plant_Code"))

    '        Rep.Execute()
    '        If Rep.Success Then
    '            If Rep.ErrMessage = Nothing Then
    '                POs = Rep.Data
    '                If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0219")
    '                End If
    '                If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '                    POs.Columns.Remove("EKPO-ZWERT")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0218")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0220")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '                    POs.Columns.Remove("EKKO-MEMORYTYPE")
    '                End If

    '                Dim TN As New DataColumn
    '                Dim SB As New DataColumn

    '                TN.ColumnName = "Usuario"
    '                TN.Caption = "Usuario"
    '                TN.DefaultValue = "BM4691"

    '                SB.DefaultValue = SAP
    '                SB.ColumnName = "SAPBox"
    '                SB.Caption = "SAPBox"

    '                POs.Columns.Add(TN)
    '                POs.Columns.Add(SB)

    '                cn.AppendTableToSqlServer("SC_OpenOrders", POs)
    '            Else
    '                Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '                BGOI_L6P.ReportProgress(0)
    '            End If
    '        Else
    '            Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '            BGOI_L6P.ReportProgress(0)
    '        End If


    '        ''******************************************************************************
    '        ''* Las POs GR103 ya no deben mostarse en el reporte al formar parte de ITrade *
    '        ''******************************************************************************
    '        ''Dim Rep2 As New SAPCOM.OpenGR105_Report(SAP, "BM4691", "LAT")
    '        ''Status = "Getting: " & SAP & " - GR 103 Plant: " & row("Plant_Code")

    '        ''BGOI_L6P.ReportProgress(0)

    '        ''Rep2.RepairsLevel = IncludeRepairs
    '        ''Rep2.Include_GR_IR = True
    '        ''Rep2.IncludeDelivDates = True
    '        ''Rep2.IncludePlant(row("Plant_Code"))

    '        ''Rep2.Execute()
    '        ''If Rep2.Success Then
    '        ''    If Rep2.ErrMessage = Nothing Then
    '        ''        POs = Rep2.Data
    '        ''        If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '        ''            POs.Columns.Remove("EKKO-WAERS-0219")
    '        ''        End If
    '        ''        If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '        ''            POs.Columns.Remove("EKPO-ZWERT")
    '        ''        End If
    '        ''        If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '        ''            POs.Columns.Remove("EKKO-WAERS-0218")
    '        ''        End If
    '        ''        If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '        ''            POs.Columns.Remove("EKKO-WAERS-0220")
    '        ''        End If
    '        ''        If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '        ''            POs.Columns.Remove("EKKO-MEMORYTYPE")
    '        ''        End If

    '        ''        Dim TN2 As New DataColumn
    '        ''        Dim SB2 As New DataColumn

    '        ''        TN2.ColumnName = "Usuario"
    '        ''        TN2.Caption = "Usuario"
    '        ''        TN2.DefaultValue = "BM4691"

    '        ''        SB2.DefaultValue = SAP
    '        ''        SB2.ColumnName = "SAPBox"
    '        ''        SB2.Caption = "SAPBox"

    '        ''        POs.Columns.Add(TN2)
    '        ''        POs.Columns.Add(SB2)

    '        ''        Try
    '        ''            cn.AppendTableToSqlServer("SC_OpenOrders", POs)
    '        ''        Catch ex As Exception
    '        ''            InsertRowByRow(POs)
    '        ''        End Try

    '        ''    End If
    '        ''End If
    '    Next
    'End Sub
    'Private Sub BGOI_L6P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOI_L6P.ProgressChanged
    '    lblStatus.Text = Status
    '    lstStatus.Items.Add(Status)
    'End Sub
    'Private Sub BGOI_L6P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOI_L6P.RunWorkerCompleted
    '    EndProcess()
    'End Sub
    'Private Sub BGOI_G4P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOI_G4P.DoWork
    '    Dim SAP As String = "G4P"
    '    Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)

    '    Dim POs As New DataTable
    '    Dim cn As New OAConnection.Connection
    '    dtPlants.Tabla = "SC_Plant"
    '    dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
    '    dtPlants.Execute()

    '    For Each row As DataRow In dtPlants.Data.Rows
    '        Dim Rep As New SAPCOM.OpenOrders_Report(SAP, "BM4691", "LAT")
    '        Status = "Getting: " & SAP & " - Plant: " & row("Plant_Code")

    '        BGOI_G4P.ReportProgress(0)

    '        Rep.RepairsLevel = IncludeRepairs
    '        Rep.Include_GR_IR = True
    '        Rep.IncludeDelivDates = True
    '        Rep.IncludePlant(row("Plant_Code"))

    '        Rep.Execute()
    '        If Rep.Success Then
    '            If Rep.ErrMessage = Nothing Then
    '                POs = Rep.Data
    '                If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0219")
    '                End If
    '                If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
    '                    POs.Columns.Remove("EKPO-ZWERT")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0218")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
    '                    POs.Columns.Remove("EKKO-WAERS-0220")
    '                End If
    '                If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
    '                    POs.Columns.Remove("EKKO-MEMORYTYPE")
    '                End If

    '                Dim TN As New DataColumn
    '                Dim SB As New DataColumn

    '                TN.ColumnName = "Usuario"
    '                TN.Caption = "Usuario"
    '                TN.DefaultValue = "BM4691"

    '                SB.DefaultValue = SAP
    '                SB.ColumnName = "SAPBox"
    '                SB.Caption = "SAPBox"

    '                POs.Columns.Add(TN)
    '                POs.Columns.Add(SB)

    '                cn.AppendTableToSqlServer("SC_OpenOrders", POs)
    '            Else
    '                Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '                BGOI_G4P.ReportProgress(0)
    '            End If
    '        Else
    '            Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
    '            BGOI_G4P.ReportProgress(0)
    '        End If

    '    Next
    'End Sub
    'Private Sub BGOI_G4P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOI_G4P.ProgressChanged
    '    lblStatus.Text = Status
    '    lstStatus.Items.Add(Status)
    'End Sub
    'Private Sub BGOI_G4P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOI_G4P.RunWorkerCompleted
    '    EndProcess()
    'End Sub
    Private Sub InsertRowByRow(ByVal T As DataTable)
        Dim T2 As New DataTable
        Dim cn As New OAConnection.Connection

        T2 = T.Clone

        Try
            For Each X As DataRow In T.Rows
                T2.Clear()
                T2.ImportRow(X)
                cn.AppendTableToSqlServer("SC_OpenOrders", T2)
            Next

        Catch ex As Exception
            'Do nothing
        End Try
    End Sub
    Public Function GetOwner(ByVal pSAP As String, Optional ByVal pSpend As String = Nothing, Optional ByVal pPlant As String = Nothing, Optional ByVal pPGrp As String = Nothing, Optional ByVal pPOrg As String = Nothing) As Owner
        Dim cn As New OAConnection.Connection
        Dim Data As DataTable
        Dim Where As String = ""

        Try
            If Not pSpend Is Nothing Then
                Where = "(([Spend] = 0) or ([Spend] = " & pSpend & "))"
            End If

            If Not pPlant Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                Where = Where & "((Plant = '') or (Plant = '" & pPlant & "'))"
            End If

            If Not pPGrp Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                Where = Where & "(([P Grp] = '') or ([P Grp] = '" & pPGrp & "'))"
            End If

            If Not pPOrg Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                Where = Where & "(([P Org] = '') or ([P Org] = '" & pPOrg & "'))"
            End If


            Data = cn.RunSentence("Select *,0 as Value From LA_Indirect_Distribution Where (SAPBox = '" & pSAP & "')" & IIf(Where <> "", " And (" & Where & ")", "")).Tables(0)
            If Data.Rows.Count > 0 Then
                If Data.Rows.Count = 1 Then
                    Dim T As New Owner

                    T.SPS = Data.Rows(0).Item("SPS")
                    T.Owner = Data.Rows(0).Item("Owner")
                    Return T
                Else

                    For Each r As DataRow In Data.Rows
                        Dim val As Integer = 0

                        If (r("SAPBox") = pSAP) Then
                            val += 2
                        Else
                            If r("SAPBox") = "" Then
                                val += 1
                            End If
                        End If


                        If (r("Plant") = pPlant) Then
                            val += 2
                        Else
                            If r("Plant") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Org") = pPOrg) Then
                            val += 2
                        Else
                            If r("P Org") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Grp") = pPGrp) Then
                            val += 2
                        Else
                            If r("P Grp") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("Spend") = pSpend) Then
                            val += 2
                        Else
                            If r("Spend") = 0 Then
                                val += 1
                            End If
                        End If

                        r("Value") = val
                    Next

                    Dim T As New Owner
                    Dim SPS = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("SPS")).ToList()
                    Dim DOwner = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("Owner")).ToList()

                    T.SPS = SPS(0)
                    T.Owner = DOwner(0)

                    'MsgBox("Multiple choises for:" & Chr(13) & Chr(13) & "SAPBox: " & pSAP & Chr(13) & "LE: " & pLE & Chr(13) & "Plant:" & pPlant & Chr(13) & "Vendor: " & pVendor & Chr(13) & "Mat. Grp: " & pMatGrp)
                    Return T
                End If
            Else
                ' MsgBox("Rules can't be found")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Sub MSG(ByVal MSG As String)
        ' lstStatus.Items.Add(MSG)
    End Sub
    Private Sub GetReports()
        ' X = New Manager
        X.MyEvent = [Delegate].Combine(X.MyEvent, New Manager.EventFirm(AddressOf Me.MSG))

        X.MyFirmEvent = [Delegate].Combine(X.MyFirmEvent, New Manager.MyFirm(AddressOf Me.YourMessage))

        pgrWorking.Style = ProgressBarStyle.Marquee
        MSG("Start: [" & Now.ToString & "]")
        'lstStatus.Items.Add("Start: [" & Now.ToString & "]")

        Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)
        dtPlants.Tabla = "SC_Plant"
        dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
        dtPlants.Execute()

        For Each row As DataRow In dtPlants.Data.Rows
            Dim WL7 As New OpenOrderWorker
            WL7.Plant = row("Plant_Code")
            WL7.SAPBox = "L7P"
            WL7.EventoAPublicar = [Delegate].Combine(WL7.EventoAPublicar, New OpenOrderWorker.FirmaEventoAPublicar(AddressOf X.AvisemeAqui))
            X.AddWorker(WL7)

            Dim WL6 As New OpenOrderWorker
            WL6.Plant = row("Plant_Code")
            WL6.SAPBox = "L6P"
            WL6.EventoAPublicar = [Delegate].Combine(WL6.EventoAPublicar, New OpenOrderWorker.FirmaEventoAPublicar(AddressOf X.AvisemeAqui))
            X.AddWorker(WL6)

            Dim WG4 As New OpenOrderWorker
            WG4.Plant = row("Plant_Code")
            WG4.SAPBox = "G4P"
            WG4.EventoAPublicar = [Delegate].Combine(WG4.EventoAPublicar, New OpenOrderWorker.FirmaEventoAPublicar(AddressOf X.AvisemeAqui))
            X.AddWorker(WG4)

            Dim WBG As New OpenOrderWorker
            WBG.Plant = row("Plant_Code")
            WBG.SAPBox = "GBP"
            WBG.EventoAPublicar = [Delegate].Combine(WBG.EventoAPublicar, New OpenOrderWorker.FirmaEventoAPublicar(AddressOf X.AvisemeAqui))
            X.AddWorker(WBG)
        Next

        X.RunWorkers()
    End Sub
    'Public Sub F() Handles X.I_Finished
    '    'Writemsg("End: [" & Now.Date.ToString & "]")
    '    MsgBox("End: [" & Now.ToString & "]")
    '    ' RaiseEvent Inf()
    'End Sub  
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.lblBGW.Text = "Finished: " & X.Total_Finished & " of " & X.Workers
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        X = New Manager
        X.MyEvent = [Delegate].Combine(X.MyEvent, New Manager.EventFirm(AddressOf Me.MSG))

        pgrWorking.Style = ProgressBarStyle.Marquee
        MSG("Start: [" & Now.ToString & "]")
        'lstStatus.Items.Add("Start: [" & Now.ToString & "]")

        'Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)
        'dtPlants.Tabla = "SC_Plant"
        'dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
        'dtPlants.Execute()


        Dim WL7 As New OpenOrderWorker
        WL7.Plant = ("0301")
        WL7.SAPBox = "L7P"
        WL7.EventoAPublicar = [Delegate].Combine(WL7.EventoAPublicar, New OpenOrderWorker.FirmaEventoAPublicar(AddressOf X.AvisemeAqui))
        X.AddWorker(WL7)

        X.RunWorkers()
    End Sub
    Private Sub BGW_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGW.RunWorkerCompleted
        MsgBox("Finish")
    End Sub


    '***************************************************************************************

    Public Sub YourMessage(ByVal Message As String)
        RaiseEvent M(Message)
    End Sub

    Private Sub Restart_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Restart.Tick
        lblStatus.Text = "Waiting: " & res & " to restart."
        res -= 1
        If res < 0 Then
            System.Diagnostics.Process.Start("OpenItemKiller.exe")
        End If
    End Sub

End Class

Public Class Owner
    Public SPS
    Public Owner
End Class

Public Class OpenOrderWorker
    Friend WithEvents BG As System.ComponentModel.BackgroundWorker
    Public Delegate Sub FirmaEventoAPublicar(ByVal SAPBox As String, ByVal Plant As String, ByVal OO As DataTable, ByVal CNFEKES As DataTable, ByVal CNFEKKO As DataTable, ByVal NAST As DataTable, ByVal ErMsg As List(Of String))
    Public EventoAPublicar As FirmaEventoAPublicar

    Event Done()
    Private _SAPBox As String = Nothing
    Private _Plant As String = Nothing
    Private _Finished As Boolean
    Private _Status As String

    Public Property Finished() As Boolean
        Get
            Return _Finished
        End Get
        Set(ByVal value As Boolean)
            _Finished = value
        End Set
    End Property
    Public Property SAPBox() As String
        Get
            Return _SAPBox
        End Get
        Set(ByVal value As String)
            _SAPBox = value
        End Set
    End Property
    Public Property Plant() As String
        Get
            Return _Plant
        End Get
        Set(ByVal value As String)
            _Plant = value
        End Set
    End Property
    Private Sub MyWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BG.DoWork
        Dim T As Integer = 0
        Dim SC As New SAPCOM.SAPConnector
        Dim ErMsg As New List(Of String)

The_Process:
        Dim C As New Object
        C = SC.GetSAPConnection(_SAPBox, "BM4691", "LAT")

        'Dim OO As New DataTable
        'Dim Cnf As New DataTable

        'Dim dtPlants As New OAConnection.SQLInstruction(eSQLInstruction.Select)
        Dim POs As New DataTable
        'Dim cn As New OAConnection.Connection
        Dim GetConfirmation As Boolean
        'dtPlants.Tabla = "SC_Plant"
        'dtPlants.AgregarParametro(New SQLInstrucParam("Plant_Code", "", False))
        'dtPlants.Execute()

        GetConfirmation = False
        'Dim Rep As New   SAPCOM.OpenOrders_Report(_SAPBox, "BM4691", "LAT")
        If C Is Nothing Then
            If T <= 3 Then
                GoTo The_Process
                T += 1
            Else
                ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message: Couldn't get SAP connection.")
                GoTo The_End
            End If
        End If

        Dim Rep As New SAPCOM.OpenOrders_Report(C)

        ' _Status = "[" & Now.Date.ToString & "] Getting: " & _SAPBox & " - Plant: " & _Plant
        'BGOI_L7P.ReportProgress(0)

        Rep.RepairsLevel = IncludeRepairs
        Rep.Include_GR_IR = True
        Rep.IncludeDelivDates = True
        Rep.Include_YO_Ref = True
        Rep.IncludePlant(_Plant)

        Rep.ExcludeMatGroup("S731516AW")
        Rep.ExcludeMatGroup("S801416AQ")
        Rep.ExcludeMatGroup("S731516AV")

        Rep.Execute()
        'Dim EKES As New SAPCOM.EKES_Report(_SAPBox, "BM4691", "LAT")
        'Dim EKKO As New SAPCOM.EKKO_Report(_SAPBox, "BM4691", "LAT")

        Dim EKES As New SAPCOM.EKES_Report(C)
        Dim EKKO As New SAPCOM.EKKO_Report(C)
        Dim NAST As New SAPCOM.NAST_Report(C)

        If Rep.Success Then
            If Rep.ErrMessage = Nothing Then
                POs = Rep.Data
                GetConfirmation = True
                If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
                    POs.Columns.Remove("EKKO-WAERS-0219")
                End If
                If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
                    POs.Columns.Remove("EKPO-ZWERT")
                End If
                If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
                    POs.Columns.Remove("EKKO-WAERS-0218")
                End If
                If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
                    POs.Columns.Remove("EKKO-WAERS-0220")
                End If
                If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
                    POs.Columns.Remove("EKKO-MEMORYTYPE")
                End If

                Dim TN As New DataColumn
                Dim SB As New DataColumn

                TN.ColumnName = "Usuario"
                TN.Caption = "Usuario"
                TN.DefaultValue = "BM4691"

                SB.DefaultValue = _SAPBox
                SB.ColumnName = "SAPBox"
                SB.Caption = "SAPBox"

                POs.Columns.Add(TN)
                POs.Columns.Add(SB)

                'cn.AppendTableToSqlServer("SC_OpenOrders", POs)
                'OO.Merge(POs)

                Dim cRow As DataRow
                For Each cRow In POs.Rows
                    EKKO.IncludeDocument(cRow.Item("Doc Number"))
                    EKES.IncludeDocument(cRow.Item("Doc Number"))
                    NAST.IncludeDocument(cRow.Item("Doc Number"))
                Next
            Else
                'Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
                'BGOI_L7P.ReportProgress(0)
            End If

            If GetConfirmation Then
                'Status = "Getting: " & SAP & " - Confirmation: " & row("Plant_Code")
                'BGOI_L7P.ReportProgress(0)

                'Status = "Getting: " & SAP & " - EKES: " & row("Plant_Code")
                'BGOI_L7P.ReportProgress(0)

                EKES.Execute()
                If EKES.Success Then
                    Dim SBE As New DataColumn
                    SBE.DefaultValue = _SAPBox
                    SBE.ColumnName = "SAPBox"
                    SBE.Caption = "SAPBox"
                    EKES.Data.Columns.Add(SBE)

                    If EKES.Data.Columns.IndexOf("OA") <> -1 Then
                        EKES.Data.Columns.Remove("OA")
                    End If

                    If EKES.Data.Columns.IndexOf("O Reference") <> -1 Then
                        EKES.Data.Columns.Remove("O Reference")
                    End If

                    'cn.AppendTableToSqlServer("DMS_Confirmation", EKES.Data)
                    'Cnf.Merge(EKES.Data)

                    'cn.ExecuteInServer("Delete From DMS_Confirmation Where Confirmation = 'YV'")
                Else
                    ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & EKES.ErrMessage)
                    'Status = "-> Fail: " & SAP & " - EKES: " & row("Plant_Code") & "[" & EKES.ErrMessage & "]"
                    'BGOI_L7P.ReportProgress(0)
                End If

                'Status = "Getting: " & SAP & " - EKKO: " & row("Plant_Code")
                'BGOI_L7P.ReportProgress(0)

                EKKO.Execute()
                If EKKO.Success Then
                    Dim ESB As New DataColumn
                    ESB.DefaultValue = _SAPBox
                    ESB.ColumnName = "SAPBox"
                    ESB.Caption = "SAPBox"

                    If EKKO.Data.Columns.IndexOf("Company Code") <> -1 Then
                        EKKO.Data.Columns.Remove("Company Code")
                    End If
                    If EKKO.Data.Columns.IndexOf("Doc Type") <> -1 Then
                        EKKO.Data.Columns.Remove("Doc Type")
                    End If
                    If EKKO.Data.Columns.IndexOf("Created On") <> -1 Then
                        EKKO.Data.Columns.Remove("Created On")
                    End If
                    If EKKO.Data.Columns.IndexOf("Created By") <> -1 Then
                        EKKO.Data.Columns.Remove("Created By")
                    End If
                    If EKKO.Data.Columns.IndexOf("Vendor") <> -1 Then
                        EKKO.Data.Columns.Remove("Vendor")
                    End If
                    If EKKO.Data.Columns.IndexOf("Language") <> -1 Then
                        EKKO.Data.Columns.Remove("Language")
                    End If
                    If EKKO.Data.Columns.IndexOf("POrg") <> -1 Then
                        EKKO.Data.Columns.Remove("POrg")
                    End If
                    If EKKO.Data.Columns.IndexOf("PGrp") <> -1 Then
                        EKKO.Data.Columns.Remove("PGrp")
                    End If
                    If EKKO.Data.Columns.IndexOf("Currency") <> -1 Then
                        EKKO.Data.Columns.Remove("Currency")
                    End If
                    If EKKO.Data.Columns.IndexOf("Doc Date") <> -1 Then
                        EKKO.Data.Columns.Remove("Doc Date")
                    End If
                    If EKKO.Data.Columns.IndexOf("Validity Start") <> -1 Then
                        EKKO.Data.Columns.Remove("Validity Start")
                    End If
                    If EKKO.Data.Columns.IndexOf("Validity End") <> -1 Then
                        EKKO.Data.Columns.Remove("Validity End")
                    End If
                    If EKKO.Data.Columns.IndexOf("Y Refer") <> -1 Then
                        EKKO.Data.Columns.Remove("Y Refer")
                    End If
                    If EKKO.Data.Columns.IndexOf("SalesPerson") <> -1 Then
                        EKKO.Data.Columns.Remove("SalesPerson")
                    End If
                    If EKKO.Data.Columns.IndexOf("Telephone") <> -1 Then
                        EKKO.Data.Columns.Remove("Telephone")
                    End If

                    EKKO.Data.Columns.Add(ESB)
                    For Each r As DataRow In EKKO.Data.Rows
                        If r("O Reference").ToString.ToUpper.IndexOf("Y") <> -1 Then
                            r("O Reference") = "Manual"
                        Else
                            r("O Reference") = ""
                        End If
                    Next

                    NAST.Show_All_Records = True
                    NAST.AddCustomField("AENDE", "Chance")

                    If NAST.IsReady Then
                        NAST.Execute()
                        If NAST.Success Then
                            Dim NSB As New DataColumn
                            NSB.DefaultValue = _SAPBox
                            NSB.ColumnName = "SAPBox"
                            NSB.Caption = "SAPBox"
                            NAST.Data.Columns.Add(NSB)
                        Else
                            ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & NAST.ErrMessage)
                        End If
                    End If

                    'cn.AppendTableToSqlServer("DMS_Confirmation", EKKO.Data)
                    'Cnf.Merge(EKKO.Data)
                    'cn.ExecuteInServer("Delete From DMS_Confirmation Where Confirmation = ''")
                Else
                    ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & EKKO.ErrMessage)
                End If
            End If
        Else
            ErMsg.Add(Now.Date.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message:" & Rep.ErrMessage)
            'Status = "-> Fail: " & SAP & " - Plant: " & row("Plant_Code") & " :: " & Rep.ErrMessage
            'BGOI_L7P.ReportProgress(0)
        End If

        '****************************************************************
        'Las POs GR103 ya no deben mostarse en el reporte al formar parte de ITrade

        'Dim Rep2 As New SAPCOM.OpenGR105_Report(SAP, "BM4691", "LAT")
        'Status = "Getting: " & SAP & " - GR 103 Plant: " & row("Plant_Code")
        'BGOI_L7P.ReportProgress(0)

        'Rep2.RepairsLevel = IncludeRepairs
        'Rep2.Include_GR_IR = True
        'Rep2.IncludeDelivDates = True
        'Rep2.IncludePlant(row("Plant_Code"))

        'Rep2.Execute()
        'If Rep2.Success Then
        '    If Rep2.ErrMessage = Nothing Then
        '        POs = Rep2.Data
        '        GetConfirmation = True

        '        If POs.Columns.IndexOf("EKKO-WAERS-0219") <> -1 Then
        '            POs.Columns.Remove("EKKO-WAERS-0219")
        '        End If
        '        If POs.Columns.IndexOf("EKPO-ZWERT") <> -1 Then
        '            POs.Columns.Remove("EKPO-ZWERT")
        '        End If
        '        If POs.Columns.IndexOf("EKKO-WAERS-0218") <> -1 Then
        '            POs.Columns.Remove("EKKO-WAERS-0218")
        '        End If
        '        If POs.Columns.IndexOf("EKKO-WAERS-0220") <> -1 Then
        '            POs.Columns.Remove("EKKO-WAERS-0220")
        '        End If
        '        If POs.Columns.IndexOf("EKKO-MEMORYTYPE") <> -1 Then
        '            POs.Columns.Remove("EKKO-MEMORYTYPE")
        '        End If

        '        Dim cRow As DataRow
        '        For Each cRow In Rep2.Data.Rows
        '            EKKO.IncludeDocument(cRow.Item("Doc Number"))
        '        Next

        '        For Each cRow In Rep2.Data.Rows
        '            EKES.IncludeDocument(cRow.Item("Doc Number"))
        '        Next

        '        Dim TN2 As New DataColumn
        '        Dim SB2 As New DataColumn

        '        TN2.ColumnName = "Usuario"
        '        TN2.Caption = "Usuario"
        '        TN2.DefaultValue = "BM4691"

        '        SB2.DefaultValue = SAP
        '        SB2.ColumnName = "SAPBox"
        '        SB2.Caption = "SAPBox"

        '        POs.Columns.Add(TN2)
        '        POs.Columns.Add(SB2)

        '        Try
        '            cn.AppendTableToSqlServer("SC_OpenOrders", POs)
        '        Catch ex As Exception
        '            InsertRowByRow(POs)
        '        End Try
        '    End If
        'End If

        
        'Next

        'cn.AppendTableToSqlServer("SC_OpenOrders", POs)
The_End:
        ErMsg.Add(Now.ToString & "SAP: " & _SAPBox & " Plant: " & _Plant & " Message: Finished.")
        OcurrioEvento(_SAPBox, _Plant, POs, EKES.Data, EKKO.Data, NAST.Data, ErMsg)
        _Finished = True
    End Sub
    Public Sub DoYourWork()
        BG.RunWorkerAsync()
    End Sub
    Public Sub New()
        BG = New System.ComponentModel.BackgroundWorker
        BG.WorkerReportsProgress = True
    End Sub
    Public Sub New(ByVal SAPbox As String, ByVal Plant As String, ByVal OO As DataTable, ByVal CNFEKES As DataTable, ByVal CNFEKKO As DataTable)
        BG = New System.ComponentModel.BackgroundWorker
        BG.WorkerReportsProgress = True
        _SAPBox = SAPbox
        _Plant = Plant
    End Sub
    Public Sub OcurrioEvento(ByVal SAPBox As String, ByVal Plant As String, ByVal OO As DataTable, ByVal CNFEKES As DataTable, ByVal CNFEKKO As DataTable, ByVal Nast As DataTable, ByVal ErMsg As List(Of String))
        EventoAPublicar(SAPBox, Plant, OO, CNFEKES, CNFEKKO, NAST, ErMsg)
    End Sub
End Class

Public Class Manager
    '************************************
    Public Delegate Sub MyFirm(ByVal Message As String)
    Public MyFirmEvent As MyFirm

    Public Sub RaseMyEvent(ByVal Message As String)
        MyFirmEvent(Message)
    End Sub

    '************************************
    Friend WithEvents BG As System.ComponentModel.BackgroundWorker
    Public Delegate Sub EventFirm(ByVal MSG As String)
    Public MyEvent As EventFirm
    Friend Event I_Finished()
    Private _WorkerList As New List(Of OpenOrderWorker)
    Private _Finished As Integer = 0
    Private _OOR As New DataTable
    Private _CNFEKES As New DataTable
    Private _CNFEKKO As New DataTable
    Private _NAST As New DataTable
    Private _OO_List As New List(Of DataTable)
    Private _CNFEKES_List As New List(Of DataTable)
    Private _CNFEKKO_List As New List(Of DataTable)
    Private _NAST_List As New List(Of DataTable)
    Private _ErMsg As New List(Of String)
    Private texto As String = ""
    Event MyReport(ByVal Message As String)
    Private _D As DataTable

    '*******************************************************************

    Public ReadOnly Property D() As DataTable
        Get
            Return _D
        End Get
    End Property
    Public ReadOnly Property Total_Finished() As Integer
        Get
            Return _Finished
        End Get
    End Property
    Public ReadOnly Property Workers() As Integer
        Get
            Return _WorkerList.Count
        End Get
    End Property
    Public Sub AvisemeAqui(ByVal SAPBox As String, ByVal Plant As String, ByVal OO As DataTable, ByVal CNFEKES As DataTable, ByVal CNFEKKO As DataTable, ByVal NAST As DataTable, ByVal ErMsg As List(Of String))

        ' RaseMyEvent(SAPBox & "-" & Plant)
        'RaiseEvent MyReport(SAPBox & "-" & Plant)

        Dim r As DataRow = _D.NewRow

        Dim EL As String = SAPBox & "-" & Plant

        For Each e In ErMsg
            EL = EL & " / " & e
        Next

        r("Message") = EL
        _D.Rows.Add(r)

        Dim Finish As Boolean = True
        Dim T As Integer = _WorkerList.Count

        _Finished = _Finished + 1

        If Not _OOR Is Nothing Then
            _OO_List.Add(OO)

        End If
        If Not CNFEKES Is Nothing Then
            _CNFEKES_List.Add(CNFEKES)

        End If
        If Not CNFEKKO Is Nothing Then
            _CNFEKKO_List.Add(CNFEKKO)
        End If
        If Not ErMsg Is Nothing Then
            For Each E In ErMsg
                texto = texto & Chr(13) & E
            Next
        End If

        If Not NAST Is Nothing Then
            _NAST_List.Add(NAST)
        End If

        If T = _Finished Then
            Dim File As String

            If Not My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\DownloadLog.txt") Then
                Dim fic As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\DownloadLog.txt"
                Dim texto As String = ""

                Dim scw As New System.IO.StreamWriter(fic)
                scw.WriteLine(texto)
                scw.Close()
            End If

            File = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\DownloadLog.txt"
            Dim sw As New System.IO.StreamWriter(File, True)
            sw.WriteLine(texto)
            sw.Close()

            Dim cn As New OAConnection.Connection

            For Each DT As DataTable In _OO_List
                _OOR.Merge(DT)
            Next

            For Each DT As DataTable In _CNFEKES_List
                _CNFEKES.Merge(DT)
            Next

            For Each dt As DataTable In _CNFEKKO_List
                _CNFEKKO.Merge(dt)
            Next

            For Each DT As DataTable In _NAST_List
                _NAST.Merge(DT)
            Next

            cn.ExecuteInServer("Delete From SC_OpenOrders")
            cn.ExecuteInServer("Delete From DMS_Confirmation")
            cn.ExecuteInServer("Delete From [LA_Transmition_NAST]")

            cn.AppendTableToSqlServer("SC_OpenOrders", _OOR)
            cn.AppendTableToSqlServer("DMS_Confirmation", _CNFEKES)
            cn.AppendTableToSqlServer("DMS_Confirmation", _CNFEKKO)
            cn.AppendTableToSqlServer("LA_Transmition_NAST", _NAST)

            cn.ExecuteInServer("Delete From DMS_Confirmation Where (Confirmation = '')")
            RaiseEvent I_Finished()
        End If
    End Sub
    Private Sub MyWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BG.DoWork
        For Each W As OpenOrderWorker In _WorkerList
            W.DoYourWork()
        Next
    End Sub
    Public Sub AddWorker(ByVal Worker As OpenOrderWorker)
        _WorkerList.Add(Worker)
    End Sub
    Public Sub RunWorkers()
        BG.RunWorkerAsync()
    End Sub
    Public Sub New()
        BG = New System.ComponentModel.BackgroundWorker
        BG.WorkerReportsProgress = True
        _D = New DataTable
        _D.Columns.Add("Message")
    End Sub
    Public Sub Fin() Handles Me.I_Finished
        Dim cn As New OAConnection.Connection
        cn.ExecuteInServer("DELETE FROM SC_OpenOrders WHERE ([Doc Type] <> 'EC') AND (SAPBox <> 'L7P') AND (NOT EXISTS (SELECT Name, Tnumber, Status FROM dbo.[PSS People] WHERE (dbo.SC_OpenOrders.[Created By] = Tnumber)))")
        cn.ExecuteInServer("DELETE FROM SC_OpenOrders WHERE ([Doc Type] = 'NB') AND (SAPBox = 'L7P') AND (NOT EXISTS (SELECT Name, Tnumber, Status FROM dbo.[PSS People] WHERE (dbo.SC_OpenOrders.[Created By] = Tnumber)))")
        cn.ExecuteInServer("DELETE From SC_OpenOrders Where (Vendor = '15145463')")

        'To Do:
        'Eliminar las PO de la distribucion temporal que ya no estan abiertas
        cn.ExecuteInServer("DELETE FROM dbo.LA_TMP_Open_Orders_Distribution Where (NOT EXISTS (SELECT [Doc Number] FROM SC_OpenOrders Where (dbo.LA_TMP_Open_Orders_Distribution.[Doc Number] = [Doc Number]) AND (dbo.LA_TMP_Open_Orders_Distribution.SAPBox = SAPBox)))")

        'Agregar las POs que son nuevas a la distribucion temporal
        cn.ExecuteInServer("Insert Into LA_TMP_Open_Orders_Distribution (SAPBox, [Doc Number]) SELECT DISTINCT SAPBox, [Doc Number] From SC_OpenOrders WHERE (NOT EXISTS (SELECT [Doc Number] From dbo.LA_TMP_Open_Orders_Distribution Where (dbo.SC_OpenOrders.[Doc Number] = [Doc Number]) AND (dbo.SC_OpenOrders.SAPBox = SAPBox)))")

        'Crear una funcion para asignarles el owner a las nuevas.
        Dim OO As New DataTable
        OO = cn.RunSentence("Select * From vst_LA_Check_Distribution").Tables(0)

        If OO.Rows.Count > 0 Then
            For Each r As DataRow In OO.Rows
                Try
                    Dim OI As New OAConnection.DMS_User(r("SAPBox"), r("Mat Group"), r("Purch Grp"), r("Purch Org"), r("Plant"))
                    OI.Execute()

                    If OI.Success Then
                        cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = '" & OI.SPS & "', Owner = '" & OI.PTB & "' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    Else
                        cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = 'BB0898', Owner = 'BB0898' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    End If

                    'Dim lO As New Owner
                    'lO = GetOwner(r("SAPBox"), r("Spend"), r("Plant"), r("Purch Grp"), r("Purch Org"))
                    'If Not lO Is Nothing Then
                    '    cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = '" & lO.SPS & "', Owner = '" & lO.Owner & "' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    'Else
                    '    'Si no tiene Owner se le asigna a Alejandra Baltodano
                    '    cn.ExecuteInServer("Update LA_TMP_Open_Orders_Distribution Set SPS = 'BB0898', Owner = 'BB0898' Where ((SAPBox = '" & r("SAPBox") & "') And ([Doc Number] = '" & r("Doc Number") & "'))")
                    'End If
                Catch ex As Exception

                End Try
            Next
        End If

        'Actualización de los vendors en la tabla VendorsG11.
        Dim Vn As New SAPCOM.LFA1_Report("G4P", "BM4691", "LAT")
        Dim NV As New DataTable 'New vendors

        NV = cn.RunSentence("Select * From vst_New_Vendors").Tables(0)
        If NV.Rows.Count > 0 Then
            For Each v In NV.Rows
                Vn.IncludeVendor(v("Vendor"))
            Next

            Vn.Execute()
            If Vn.Success Then
                For Each v In NV.Rows
                    Try
                        Dim VR = (From C In Vn.Data.AsEnumerable() _
                                  Where C.Item("Vendor") = v("Vendor") _
                                  Select C.Item("Country")).First

                        v("Country") = VR
                    Catch ex As Exception
                        'Do nothing
                    End Try
                Next

                cn.AppendTableToSqlServer("VendorsG11", NV)
            End If

        End If

        End

    End Sub
    Public Function GetOwner(ByVal pSAP As String, Optional ByVal pSpend As String = Nothing, Optional ByVal pPlant As String = Nothing, Optional ByVal pPGrp As String = Nothing, Optional ByVal pPOrg As String = Nothing) As Owner
        Dim cn As New OAConnection.Connection
        Dim Data As DataTable
        Dim Where As String = ""

        Try
            If Not pSpend Is Nothing Then
                Where = "(([Spend] = 0) or ([Spend] = " & pSpend & "))"
            End If

            If Not pPlant Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                'Where = Where & "((Plant = '') or (Plant = '" & pPlant & "'))"
                Where = Where & "((Plant = '" & pPlant & "'))"
            End If

            If Not pPGrp Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                ' Where = Where & "(([P Grp] = '" & pPGrp & "'))"
                Where = Where & "(([P Grp] = '') or ([P Grp] = '" & pPGrp & "'))"
            End If

            If Not pPOrg Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                'Where = Where & "(([P Org] = '') or ([P Org] = '" & pPOrg & "'))"
                Where = Where & "(([P Org] = '" & pPOrg & "'))"
            End If


            Data = cn.RunSentence("Select *,0 as Value From LA_Indirect_Distribution Where (SAPBox = '" & pSAP & "')" & IIf(Where <> "", " And (" & Where & ")", "")).Tables(0)
            If Data.Rows.Count > 0 Then
                If Data.Rows.Count = 1 Then
                    Dim T As New Owner

                    T.SPS = Data.Rows(0).Item("SPS")
                    T.Owner = Data.Rows(0).Item("Owner")
                    Return T
                Else

                    For Each r As DataRow In Data.Rows
                        Dim val As Integer = 0

                        If (r("SAPBox") = pSAP) Then
                            val += 2
                        Else
                            If r("SAPBox") = "" Then
                                val += 1
                            End If
                        End If


                        If (r("Plant") = pPlant) Then
                            val += 2
                        Else
                            If r("Plant") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Org") = pPOrg) Then
                            val += 2
                        Else
                            If r("P Org") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Grp") = pPGrp) Then
                            val += 2
                        Else
                            If r("P Grp") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("Spend") = pSpend) Then
                            val += 2
                        Else
                            If r("Spend") = 0 Then
                                val += 1
                            End If
                        End If

                        r("Value") = val
                    Next

                    Dim T As New Owner
                    Dim SPS = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("SPS")).ToList()
                    Dim DOwner = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("Owner")).ToList()

                    T.SPS = SPS(0)
                    T.Owner = DOwner(0)

                    'MsgBox("Multiple choises for:" & Chr(13) & Chr(13) & "SAPBox: " & pSAP & Chr(13) & "LE: " & pLE & Chr(13) & "Plant:" & pPlant & Chr(13) & "Vendor: " & pVendor & Chr(13) & "Mat. Grp: " & pMatGrp)
                    Return T
                End If
            Else
                ' MsgBox("Rules can't be found")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
End Class

