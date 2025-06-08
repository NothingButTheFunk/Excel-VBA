Attribute VB_Name = "Modul2"
Sub InsertLaTeXEquation()
    Dim a As Double, b As Double, c As Double
    Dim latexEq As String
    Dim filePath As String
    Dim pdfPath As String
    Dim imgPath As String

    ' Get values from Excel
    a = Range("C2").Value
    b = Range("C3").Value
    c = a * b
    Range("C4").Value = c

    ' Create LaTeX formula
    latexEq = "$$ c = f \cdot k = " & a & " \cdot " & b & " = " & c & " $$"

    ' Set file paths
    filePath = "C:\Temp\equation.tex"
    pdfPath = "C:\Temp\equation.pdf"
    imgPath = "C:\Temp\equation-1.png"

    ' Create LaTeX file
    Open filePath For Output As #1
    Print #1, "\documentclass{article}"
    Print #1, "\usepackage{amsmath}"
    Print #1, "\begin{document}"
    Print #1, latexEq
    Print #1, "\end{document}"
    Close #1

    ' Create PDF 
    ' Install https://miktex.org/ and install all updates
    ' The WaitProcess() function ensures that the Shell commands are been carried out correctly
    Call WaitForProcess("pdflatex -output-directory=C:\Temp " & filePath)

    ' Convert PDF to PNG with pdftoppm
    ' https://www.xpdfreader.com/download.html install the xpdfreader
    Call WaitForProcess("pdftoppm -png C:\Temp\equation.pdf C:\Temp\equation")

    'insert image, crop and re-orientate
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

    ' Get position of cell B7
    cellLeft = Range("B7").Left
    cellTop = Range("B7").Top

    ' Insert image
    Set pic = ActiveSheet.Shapes.AddPicture(imgPath, _
        msoFalse, msoTrue, cellLeft, cellTop, -1, -1)

    ' Crop image
    With pic.PictureFormat
        .CropLeft = 250
        .CropTop = 125
        .CropRight = 240
        .CropBottom = 675
    End With

    'position image
    pic.Left = cellLeft
    pic.Top = cellTop

    'MsgBox "Image successfully cropped and correctly positioned!", vbInformation
End Sub


