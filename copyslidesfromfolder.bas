Attribute VB_Name = "Module2"
Option Explicit
Sub CopylidesFromFolderS()
Dim folderpath As String
Dim filename As String
Dim file As String
Dim PowerPointApp As PowerPoint.Application
Dim myPresentation As PowerPoint.Presentation
Dim activePresentation1 As PowerPoint.Presentation
Dim activeSlideIndex As Integer
Dim i As Integer
Dim slide As PowerPoint.slide
Dim x As Integer

folderpath = "C:\Users\rg413939\OneDrive - GSK\General - RMCB Forum_Ju&QR\ADDSLIDES\"
Set activePresentation1 = activePresentation
activeSlideIndex = ActiveWindow.View.slide.SlideIndex
filename = Dir(folderpath & "*.pptx")

Do While filename <> ""
file = folderpath & filename
Set myPresentation = Presentations.Open(file)
MsgBox "Opening file: " & file
x = myPresentation.Slides.Count
    For i = 1 To x
    myPresentation.Slides(i).Copy

    activePresentation1.Slides.Paste (activeSlideIndex + 1)
    activeSlideIndex = activeSlideIndex + 1
Next i
    myPresentation.Close
    filename = Dir
Loop
End Sub
