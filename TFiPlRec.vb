Public Class TFiPlRec
    Public FIELDNAME As SAPCommon.TField
    Public VALUE As SAPCommon.TField
    Public CURRENCY As SAPCommon.TField

    Public Function setValues(pFIELDNAME As String, pVALUE As String, Optional pCURRENCY As String = "")

        FIELDNAME = New SAPCommon.TField("FIELDNAME", pFIELDNAME)
        VALUE = New SAPCommon.TField("VALUE", pVALUE)
        CURRENCY = New SAPCommon.TField("CURRENCY ", pCURRENCY)
    End Function

    Public Function getKey() As String
        Dim aKey As String
        aKey = FIELDNAME.Value
        getKey = aKey
    End Function

    Public Function getKeyR() As String
        Dim aKey As String
        aKey = FIELDNAME.Value
        getKeyR = aKey
    End Function

    Public Function toStringValue() As Object
        Dim aArray(3) As String
        aArray(0) = FIELDNAME.Value
        aArray(1) = VALUE.Value
        aArray(2) = CURRENCY.Value
        toStringValue = aArray
    End Function

End Class
