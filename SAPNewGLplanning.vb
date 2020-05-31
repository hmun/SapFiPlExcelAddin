Imports SAP.Middleware.Connector

Public Class SAPNewGLplanning
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCOPAActuals")
        End Try
    End Sub

    Public Function Post(pCOMP_CODE As String, pLEDGER As String, pFISC_YEAR As String,
                         pPERIOD_FROM As String, pPERIOD_TO As String, pVERSION As String,
                         pFields As Collection, pData As Collection, Optional pCheck As Boolean = False) As String
        Post = ""
        Try
            log.Debug("post - " & "creating Function BAPI_FAGL_PLANNING_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_FAGL_PLANNING_POST")
            log.Debug("post - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            log.Debug("post - " & "BeginContext")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat

            log.Debug("post - " & "Getting Function parameters")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oPerValue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oFieldList As IRfcTable = oRfcFunction.GetTable("FIELDLIST")
            Dim oHeaderInfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            oPerValue.Clear()
            oFieldList.Clear()
            oRETURN.Clear()
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            Else
                oRfcFunction.SetValue("TESTRUN", "")
            End If

            log.Debug("post - " & "setting header values")
            oHeaderInfo.SetValue("COMP_CODE", pCOMP_CODE)
            oHeaderInfo.SetValue("LEDGER", pLEDGER)
            oHeaderInfo.SetValue("FISC_YEAR", pFISC_YEAR)
            oHeaderInfo.SetValue("PERIOD_FROM", pPERIOD_FROM)
            oHeaderInfo.SetValue("PERIOD_TO", pPERIOD_TO)
            oHeaderInfo.SetValue("VERSION", pVERSION)

            Dim aTFiPl As TFiPl
            Dim aTFiPlRec As TFiPlRec
            Dim aFieldname As String
            Dim lCnt As Integer = 0
            log.Debug("post - " & "processing pFields")
            For Each aFieldname In pFields
                oFieldList.Append()
                oFieldList.SetValue("FIELDNAME", aFieldname)
            Next
            log.Debug("post - " & "processing pData")
            For Each aTFiPl In pData
                lCnt += 1
                oPerValue.Append()
                For Each aTFiPlRec In aTFiPl.aTFiPlCol
                    oPerValue.SetValue("POSNR", lCnt)
                    If aTFiPlRec.CURRENCY.Value <> "" Then
                        oPerValue.SetValue(aTFiPlRec.FIELDNAME.Value, CStr(Decimal.Round(CDec(aTFiPlRec.VALUE.Value), 2)))
                    Else
                        oPerValue.SetValue(aTFiPlRec.FIELDNAME.Value, aTFiPlRec.VALUE.Value)
                    End If
                Next aTFiPlRec
            Next aTFiPl
            ' call the BAPI
            log.Debug("post - " & "invoking " & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            log.Debug("post - " & "oRETURN.Count=" & CStr(oRETURN.Count))
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oRETURN.Count - 1
                Post = Post & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    log.Debug("post - " & "ErrorMessage=" & oRETURN(i).GetValue("MESSAGE"))
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Post = "Success - " & Post
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                log.Debug("post - " & "calling aSAPBapiTranctionCommit.commit()")
                aSAPBapiTranctionCommit.commit()
            End If
        Catch Ex As System.Exception
            log.Error("post - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPNewGLplanning")
            Post = "Error: Exception in SAPNewGLplanning.Post"
        Finally
            log.Debug("post - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
