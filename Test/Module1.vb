Module Module1

    Sub Main()

        Dim SC As New SAPCOM.SAPConnector
        Dim C As New Object
        C = SC.GetSAPConnection("L7P", "BM4691", "LAT")


        Dim s As New SAPCOM.EBAN_Report(C)

        Dim CNT As Boolean
        CNT = SC.TestConnection(C)
    End Sub

End Module
