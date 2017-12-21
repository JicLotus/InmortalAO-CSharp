Option Explicit On
Public Class cCola
    '                    Metodos publicos
    '
    ' Public sub Push(byval i as variant) mete el elemento i
    ' al final de la cola.
    '
    ' Public Function Pop As Variant: quita de la cola el primer elem
    ' y lo devuelve
    '
    ' Public Function VerElemento(ByVal Index As Integer) As Variant
    ' muestra el elemento numero Index de la cola sin quitarlo
    '
    ' Public Function PopByVal() As Variant: muestra el primer
    ' elemento de la cola sin quitarlo
    '
    ' Public Property Get Longitud() As Integer: devuelve la
    ' cantidad de elementos que tiene la cola.

    Private Const FRENTE = 1

    Private Cola As Collection

    Public Sub Reset()
        On Error GoTo hayerror

        Dim i As Integer

        For i = 1 To Me.Longitud
            Cola.Remove(FRENTE)
        Next i

        Exit Sub

hayerror:
        LogError("Error en Reset: " & Err.Description)


    End Sub

    Public Property Longitud() As Integer
        Get
            Return Cola.Count
        End Get
        Set(value As Integer)
        End Set
    End Property

    Private Function IndexValido(ByVal i As Integer) As Boolean
        IndexValido = i >= 1 And i <= Me.Longitud
    End Function

    Public Sub Class_Initialize()
        Cola = New Collection
    End Sub

    Public Function VerElemento(ByVal Index As Integer) As String
        On Error Resume Next
        If IndexValido(Index) Then
            'Pablo
            VerElemento = UCase$(Cola.Item(Index))
            '/Pablo
            'VerElemento = Cola(Index)
        Else
            VerElemento = 0
        End If
    End Function


    Public Sub Push(ByVal Nombre As String)
        On Error Resume Next
        'Mete elemento en la cola
        'Pablo
        Dim aux As String
        aux = DateTime.Now + " " + UCase$(Nombre)
        Call Cola.Add(aux)
        '/Pablo

        'Call Cola.Add(UCase$(Nombre))
    End Sub

    Public Function Pop() As String
        On Error Resume Next
        'Quita elemento de la cola
        If Cola.Count > 0 Then
            Pop = Cola(FRENTE)
            Call Cola.Remove(FRENTE)
        Else
            Pop = 0
        End If
    End Function

    Public Function PopByVal() As String
        On Error Resume Next
        'Call LogTarea("PopByVal SOS")

        'Quita elemento de la cola
        If Cola.Count > 0 Then
            PopByVal = Cola.Item(1)
        Else
            PopByVal = 0
        End If

    End Function

    Public Function Existe(ByVal Nombre As String) As Boolean
        On Error Resume Next

        Dim V As String
        Dim i As Integer
        Dim NombreEnMayusculas As String
        NombreEnMayusculas = UCase$(Nombre)

        For i = 1 To Me.Longitud
            'Pablo
            V = Mid(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
            '/Pablo
            'V = Me.VerElemento(i)
            If V = NombreEnMayusculas Then
                Existe = True
                Exit Function
            End If
        Next
        Existe = False

    End Function

    Public Sub Quitar(ByVal Nombre As String)
        On Error Resume Next
        Dim V As String
        Dim i As Integer
        Dim NombreEnMayusculas As String

        NombreEnMayusculas = UCase$(Nombre)

        For i = 1 To Me.Longitud
            'Pablo
            V = Mid(Me.VerElemento(i), 10, Len(Me.VerElemento(i)))
            '/Pablo
            'V = Me.VerElemento(i)
            If V = NombreEnMayusculas Then
                Call Cola.Remove(i)
                Exit Sub
            End If
        Next i

    End Sub

    Public Sub QuitarIndex(ByVal Index As Integer)
        On Error Resume Next
        If IndexValido(Index) Then Call Cola.Remove(Index)
    End Sub


    Private Sub Class_Terminate()
        'Destruimos el objeto Cola
        Cola = Nothing
    End Sub

End Class
