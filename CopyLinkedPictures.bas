Attribute VB_Name = "modCopyLinkedPictures"

Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As LongPtr
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr

Public Sub CopyLinkedPictures()

    Dim currentSlide As Slide
    For Each currentSlide In ActivePresentation.Slides

        Dim currentShape As Shape
        For Each currentShape In currentSlide.Shapes

            HandleShape currentSlide:=currentSlide, currentShape:=currentShape

        Next currentShape

    Next currentSlide

End Sub

Private Sub HandleShape(currentSlide As Slide, currentShape As Shape)

    If currentShape.Type = msoGroup Then

        Dim groupNames() As String
        ReDim Preserve groupNames(currentShape.GroupItems.Count + 1)

        Dim groupIndex As Integer
        groupIndex = 0

        Dim shapeWasCopied As Boolean
        shapeWasCopied = False

        Dim groupShape As Shape
        For Each groupShape In currentShape.GroupItems

            groupIndex = groupIndex + 1
            groupNames(groupIndex) = groupShape.Name

            If groupShape.Type = msoLinkedPicture Then

                HandleShape currentSlide:=currentSlide, currentShape:=groupShape

                shapeWasCopied = True

            End If

        Next groupShape

        If (shapeWasCopied) Then

            currentSlide.Shapes.Range(groupNames).Group

        End If
		
		Erase groupNames

    ElseIf currentShape.Type = msoLinkedPicture Then

        Dim newShape As Object
        Dim newName As String

        newName = currentShape.Name

        currentShape.Copy
        Set newShape = currentSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture, DisplayAsIcon:=msoFalse, Link:=msoFalse)

        With newShape
            .Width = currentShape.Width
            .Height = currentShape.Height
            .Left = currentShape.Left
            .Top = currentShape.Top
        End With

        currentShape.Delete

        newShape.Name = newName

        ClearClipboard

    End If

End Sub

Private Function ClearClipboard()

  OpenClipboard (0&)
  EmptyClipboard
  CloseClipboard

End Function
