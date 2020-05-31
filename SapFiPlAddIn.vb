Imports Microsoft.Office.Tools.Ribbon

Public Class SapFiPlAddIn
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private aCompanyCode As String
    Private aLedger As String
    Private aFiscalYear As String
    Private aPeriodFrom As String
    Private aPeriodTo As String
    Private aVersion As String
    Private aMaxLines As Long

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub SapFiPlAddIn_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Function getFiPlParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aSapFormat As New SAPFormat
        Dim aKey As String
        log.Debug("getFiPlParameters - " & "reading Parameter")
        aWB = Globals.SapFiPlExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP New-Gl Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap FI-Pl")
            getFiPlParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPNewGlPlanning" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPNewGlPlanning. Check if the current workbook is a valid SAP New-Gl Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap FI-Pl")
            getFiPlParameters = False
            Exit Function
        End If
        aCompanyCode = CStr(aPws.Cells(2, 2).Value)
        aLedger = CStr(aPws.Cells(3, 2).Value)
        aFiscalYear = CStr(aPws.Cells(4, 2).Value)
        aPeriodFrom = CStr(aPws.Cells(5, 2).Value)
        aPeriodTo = CStr(aPws.Cells(6, 2).Value)
        aVersion = CStr(aPws.Cells(7, 2).Value)
        aMaxLines = CLng(aPws.Cells(8, 2).Value)
        If aCompanyCode = "" Or aLedger = "" Or aFiscalYear = "" Or aPeriodFrom = "" Or aPeriodTo = "" Or aVersion = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Sap FI-Pl getFiPlParameters")
            getFiPlParameters = False
            Exit Function
        End If
        aFiscalYear = aSapFormat.unpack(aFiscalYear, 4)
        aPeriodFrom = aSapFormat.unpack(aPeriodFrom, 3)
        aPeriodTo = aSapFormat.unpack(aPeriodTo, 3)
        aVersion = aSapFormat.unpack(aVersion, 3)
        getFiPlParameters = True
    End Function

    Private Sub ButtonNewGLplanningCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNewGLplanningCheck.Click
        If checkCon() = True Then
            SAP_FiPl_exec(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNewGLplanningCheck_Click")
        End If
    End Sub

    Private Sub ButtonNewGLplanningPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonNewGLplanningPost.Click
        If checkCon() = True Then
            SAP_FiPl_exec(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonNewGLplanningPost_Click")
        End If
    End Sub

    Private Sub SAP_FiPl_exec(pTest As Boolean)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPFormat As New SAPFormat
        Dim aSAPNewGLplanning As New SAPNewGLplanning(aSapCon)
        Dim aTFiPL As New TFiPl
        Dim aData As Collection
        Dim aFieldlist As Collection
        Dim aStartLine As Long
        Dim aEndLine As Long
        Dim aLineCnt As Long

        Dim i As Long
        Dim j As Long
        Dim maxJ As Long
        Dim aRetStr As String

        Dim aFIELDNAME As String
        Dim aVALUE As Object
        Dim aCURRENCY As String

        Dim aCells As Excel.Range

        If getFiPlParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aWB = Globals.SapFiPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-PA")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SAP_FiPl_exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Globals.SapFiPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapFiPlExcelAddin.Application.EnableEvents = False
            Globals.SapFiPlExcelAddin.Application.ScreenUpdating = False
            i = 5
            ' determine the last column and create the fieldlist
            maxJ = 1
            j = 1
            aFieldlist = New Collection
            Do
                If CStr(aDws.Cells(2, maxJ).Value) = "" And CStr(aDws.Cells(1, maxJ).Value) <> "GL_ACCOUNT" Then
                    aFieldlist.Add(CStr(aDws.Cells(1, maxJ).Value))
                End If
                maxJ += 1
            Loop While CStr(aDws.Cells(1, maxJ).Value) <> ""
            aStartLine = i
            aLineCnt = 0
            aData = New Collection
            aTFiPL = New TFiPl
            j = 1
            Do
                If Left(CStr(aDws.Cells(i, maxJ).Value), 7) <> "Success" Then
                    aTFiPL = New TFiPl
                    j = 1
                    Do
                        aVALUE = ""
                        aCURRENCY = ""
                        aFIELDNAME = ""
                        If aDws.Cells(2, j).Value IsNot Nothing Then
                            aCURRENCY = CStr(aDws.Cells(2, j).Value)
                            If aDws.Cells(i, j).Value IsNot Nothing Then
                                aVALUE = FormatNumber(CDbl(aDws.Cells(i, j).Value), 2, True, False, False)
                            Else
                                aVALUE = FormatNumber(0, 2, True, False, False)
                            End If
                        Else
                            aCURRENCY = ""
                            If aDws.Cells(i, j).Value IsNot Nothing Then
                                Select Case CStr(aDws.Cells(3, j).Value)
                                    Case "DATE"
                                        Try
                                            aVALUE = CDate(aDws.Cells(i, j).Value).ToString("yyyyMMdd")
                                        Catch Exc As System.Exception
                                            aVALUE = ""
                                        End Try
                                    Case "PERIO"
                                        aVALUE = Right(aDws.Cells(i, j).Value, 4) & Left(aDws.Cells(i, j).Value, 3)
                                    Case Else
                                        If Left(aDws.Cells(3, j).Value, 1) = "U" Then
                                            aVALUE = aSAPFormat.unpack(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                        ElseIf Left(aDws.Cells(3, j).Value, 1) = "P" Then
                                            aVALUE = aSAPFormat.pspid(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                        Else
                                            aVALUE = aDws.Cells(i, j).Value
                                        End If
                                End Select
                            End If
                        End If
                        aFIELDNAME = CStr(aDws.Cells(1, j).Value)
                        aTFiPL.addTFiPl(aFIELDNAME, aVALUE, aCURRENCY)
                        j += 1
                    Loop While CStr(aDws.Cells(1, j).Value) <> ""
                    aData.Add(aTFiPL)
                    aLineCnt += 1
                    If aLineCnt >= CInt(aMaxLines) Then
                        log.Debug("SAP_FiPl_exec - " & "reached MaxLine - calling aSAPNewGLplanning.Post")
                        aEndLine = i
                        '     post the lines
                        Globals.SapFiPlExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
                        aRetStr = aSAPNewGLplanning.Post(aCompanyCode, aLedger, aFiscalYear, aPeriodFrom, aPeriodTo, aVersion, aFieldlist, aData, pCheck:=pTest)
                        aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
                        aCells.Value = aRetStr
                        aStartLine = i + 1
                        aLineCnt = 0
                        aData = New Collection
                    End If
                Else
                    aDws.Cells(i, maxJ + 1).Value = "ignored - already posted"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""
            ' post the rest
            If aData.Count > 0 Then
                log.Debug("SAP_FiPl_exec - " & "aData.Count > 0 - calling aSAPNewGLplanning.Post")
                aEndLine = i - 1
                Globals.SapFiPlExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
                aRetStr = aSAPNewGLplanning.Post(aCompanyCode, aLedger, aFiscalYear, aPeriodFrom, aPeriodTo, aVersion, aFieldlist, aData, pCheck:=pTest)
                aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
                aCells.Value = aRetStr
            End If
            log.Debug("SAP_AccDoc_execute - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapFiPlExcelAddin.Application.EnableEvents = True
            Globals.SapFiPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapFiPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapFiPlExcelAddin.Application.EnableEvents = True
            Globals.SapFiPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapFiPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SAP_AccDoc_execute failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
            log.Error("SAP_AccDoc_execute - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub


End Class
