Option Explicit On
Imports System.Text

Public Class clsSecurity

    Public Redundance As Byte

    Public Sub NAC_E_Byte(ByRef ByteArray() As Byte, ByVal code As Byte)

        Dim i As Integer 'Ponemos integer porque no manejamos paquetes más grandes de 10KB
        For i = 0 To UBound(ByteArray)
            ByteArray(i) = code Xor ByteArray(i)
        Next
    End Sub

    Public Sub NAC_D_Byte(ByRef ByteArray() As Byte, ByVal code As Byte)

        Dim i As Integer 'Ponemos integer porque no manejamos paquetes más grandes de 10KB
        For i = 0 To UBound(ByteArray)
            ByteArray(i) = ByteArray(i) Xor code
        Next
    End Sub

    Public Function NAC_E_String(ByVal t As String, ByVal code As Byte) As String

        Dim Bytes() As Byte

        Bytes = Encoding.Default.GetBytes(t)

        Call NAC_E_Byte(Bytes, code)
        'NAC_E_String = StrConv(Bytes, vbUnicode)
        NAC_E_String = Encoding.Default.GetString(Bytes)
    End Function

    Public Function NAC_D_String(ByVal t As String, ByVal code As Byte) As String

        Dim Bytes() As Byte



        'Bytes = StrConv(t, vbFromUnicode)
        Bytes = Encoding.Default.GetBytes(t)

        Call NAC_D_Byte(Bytes, code)
        'NAC_D_String = StrConv(Bytes, vbUnicode)
        NAC_D_String = Encoding.Default.GetString(Bytes)
    End Function

End Class
