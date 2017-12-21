

Imports System.Drawing.Imaging
Imports System.IO
Imports ServerAPI.Loads

Public Class Form1
    Dim NumMaps As Integer = 851
    Dim XMaxMapSize As Integer = 100
    Dim YMaxMapSize As Integer = 100
    Dim MapData(NumMaps + 1, XMaxMapSize + 1, YMaxMapSize + 1) As MapBlock

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim mapFile As String
        Dim map As Integer


        Dim numeroColumnas As Integer = 16 'hay 17 columns 0 a 16
        Dim numeroFilas As Integer = 22 '23 filas

        Dim xini = 0
        Dim yini = 0

        Dim tempMapBlock As MapBlock

        Dim tempInitialMap As Integer

        map = 743
        tempInitialMap = map

        Dim x As Integer
        Dim y As Integer

        Dim mapas(numeroColumnas, numeroFilas) As String


        Dim finalImage As System.Drawing.Bitmap
        finalImage = New System.Drawing.Bitmap((numeroColumnas + 1) * 26, (numeroFilas + 1) * 26, PixelFormat.Format24bppRgb)
        Dim g As System.Drawing.Graphics
        g = System.Drawing.Graphics.FromImage(finalImage)


        Dim imagenes(numeroColumnas, numeroFilas) As Bitmap

        Dim preMap As Integer
        Dim nombreMapa As String = ""


        For y = yini To numeroFilas
            For x = xini To numeroColumnas


                imagenes(x, y) = getBitMap(map)

                mapFile = Application.StartupPath & "\Maps\" & "Mapa" & map

                CargarMapa(MapData, map, mapFile, nombreMapa)

                tempMapBlock = MapData(map, 92, 50)
                mapas(x, y) = map.ToString() + "-" + nombreMapa

                preMap = map
                map = tempMapBlock.TileExit.map

                Dim tempY As Integer

                Dim tempMap As Integer
                tempMap = map

                If map = 0 Then

                    map = preMap
                    tempY = 2
                    While tempMap = 0
                        tempY = tempY + 1
                        tempMapBlock = MapData(map, 92, tempY)
                        tempMap = tempMapBlock.TileExit.map
                    End While
                    map = tempMap
                End If

            Next x

            mapFile = Application.StartupPath & "\Maps\" & "Mapa" & tempInitialMap
            CargarMapa(MapData, tempInitialMap, mapFile, nombreMapa)
            tempMapBlock = MapData(tempInitialMap, 50, 94)
            map = tempMapBlock.TileExit.map
            tempInitialMap = map

        Next y


        For y = yini To numeroFilas

            For x = xini To numeroColumnas

                Dim img As Bitmap
                img = imagenes(x, y)

                If (img IsNot Nothing) Then
                    g.DrawImage(img, New System.Drawing.Rectangle(x * img.Width, y * img.Height, img.Width, img.Height))
                End If
            Next x
        Next y

        finalImage.Save("ImagenMundo.bmp")


        Using outputFile As New StreamWriter(Convert.ToString("datosMundo.txt"))

            outputFile.WriteLine(numeroFilas + 1) 'netas
            outputFile.WriteLine(numeroColumnas + 1)

            For y = yini To numeroFilas
                For x = xini To numeroColumnas
                    outputFile.WriteLine(mapas(x, y).Split("-")(0))
                    outputFile.WriteLine(mapas(x, y).Split("-")(1))
                Next x
            Next y

        End Using


        MessageBox.Show("Ready!")

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Function getBitMap(num As Integer) As Bitmap
        Dim Bitmap As System.Drawing.Bitmap
        Dim a As Drawing.Size
        a = New Drawing.Size
        a.Width = 26
        a.Height = 26

        Bitmap = New System.Drawing.Bitmap("minimaps/" + num.ToString() + ".bmp")

        Dim newBitmap As Bitmap
        newBitmap = New Bitmap(Bitmap, a)

        getBitMap = newBitmap
    End Function

End Class
