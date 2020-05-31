Public Class TFiPl
    Public aTFiPlCol As Collection
    Public Sub New()
        aTFiPlCol = New Collection
    End Sub

    Public Sub addTFiPl(pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "")
        Dim aTFiPlRec As TFiPlRec
        Dim aKey As String

        aKey = pFIELDNAME
        If aTFiPlCol.Contains(aKey) Then
            aTFiPlRec = aTFiPlCol(aKey)
            aTFiPlRec.setValues(pFIELDNAME, pVALUE, pCURRENCY)
        Else
            aTFiPlRec = New TFiPlRec
            aTFiPlRec.setValues(pFIELDNAME, pVALUE, pCURRENCY)
            aTFiPlCol.Add(aTFiPlRec, aKey)
        End If
    End Sub

End Class
