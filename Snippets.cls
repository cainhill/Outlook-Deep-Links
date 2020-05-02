' http://www.cpearson.com/excel/CreateGUID.aspx

Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Function CreateGUID() As String
    Dim G As GUID
    With G
    If (CoCreateGuid(G) = 0) Then
    CreateGUID = _
        String$(8 - Len(Hex$(.Data1)), "0") & Hex$(.Data1) & _
        String$(4 - Len(Hex$(.Data2)), "0") & Hex$(.Data2) & _
        String$(4 - Len(Hex$(.Data3)), "0") & Hex$(.Data3) & _
        IIf((.Data4(0) < &H10), "0", "") & Hex$(.Data4(0)) & _
        IIf((.Data4(1) < &H10), "0", "") & Hex$(.Data4(1)) & _
        IIf((.Data4(2) < &H10), "0", "") & Hex$(.Data4(2)) & _
        IIf((.Data4(3) < &H10), "0", "") & Hex$(.Data4(3)) & _
        IIf((.Data4(4) < &H10), "0", "") & Hex$(.Data4(4)) & _
        IIf((.Data4(5) < &H10), "0", "") & Hex$(.Data4(5)) & _
        IIf((.Data4(6) < &H10), "0", "") & Hex$(.Data4(6)) & _
        IIf((.Data4(7) < &H10), "0", "") & Hex$(.Data4(7))
    End If
    End With
End Function
