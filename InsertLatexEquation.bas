Attribute VB_Name = "Modul2"
Sub InsertLaTeXEquation()
    Dim a As Double, b As Double, c As Double
    Dim latexEq As String
    Dim filePath As String
    Dim pdfPath As String
    Dim imgPath As String

    ' Werte aus Excel holen
    a = Range("C2").Value
    b = Range("C3").Value
    c = a * b
    Range("C4").Value = c

    ' LaTeX-Formel erstellen
    latexEq = "$$ c = f \cdot k = " & a & " \cdot " & b & " = " & c & " $$"

    ' Dateipfade festlegen
    filePath = "C:\Temp\equation.tex"
    pdfPath = "C:\Temp\equation.pdf"
    imgPath = "C:\Temp\equation-1.png"

    ' LaTeX-Datei erstellen
    Open filePath For Output As #1
    Print #1, "\documentclass{article}"
    Print #1, "\usepackage{amsmath}"
    Print #1, "\begin{document}"
    Print #1, latexEq
    Print #1, "\end{document}"
    Close #1

    ' PDF erzeugen
    ' Install https://miktex.org/ and install all updates
    Call WaitForProcess("pdflatex -output-directory=C:\Temp " & filePath)

    ' PDF zu PNG konvertieren mit pdftoppm
    ' https://www.xpdfreader.com/download.html install the xpdfreader
    Call WaitForProcess("pdftoppm -png C:\Temp\equation.pdf C:\Temp\equation")
    
    Call CropImage
           
End Sub

Sub WaitForProcess(cmd As String)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd.exe /c " & cmd, 1, True
End Sub


Sub CropImage()
    Dim pic As Shape
    Dim imgPath As String
    imgPath = "C:\Temp\equation-1.png"
    Dim cellLeft As Double, cellTop As Double
    Dim newWidth As Double, newHeight As Double

    ' Position von Zelle H7 berechnen
    cellLeft = Range("B7").Left
    cellTop = Range("B7").Top

    ' Bild einfügen
    Set pic = ActiveSheet.Shapes.AddPicture(imgPath, _
        msoFalse, msoTrue, cellLeft, cellTop, -1, -1)

    ' Zuschneiden auf bestimmte Bereiche (anpassen nach Bedarf)
    With pic.PictureFormat
        .CropLeft = 250
        .CropTop = 125
        .CropRight = 240
        .CropBottom = 675
    End With

    pic.Left = cellLeft
    pic.Top = cellTop

    'MsgBox "Bild erfolgreich zugeschnitten und korrekt positioniert!", vbInformation
End Sub


