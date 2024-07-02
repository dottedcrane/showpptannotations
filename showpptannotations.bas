Attribute VB_Name = "PPT_AddInns"
Option Explicit
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function

Sub Auto_Open()
      Dim NewControl As Object
      ' Store an object reference to a command bar.
      Dim ToolsMenu As CommandBars

      ' Figure out where to place the menu choice.
      Set ToolsMenu = Application.CommandBars
      
            ' Create the menu choice. The choice is created in the first position in the Tools menu.
      Set NewControl = ToolsMenu("Tools").Controls.Add(Type:=msoControlButton, Before:=1)
      With NewControl
            .DescriptionText = "Hides Objects named as Annotation SHAPE XX."
            ' Name the command.
            .Caption = "Hide Annotations"
            ' Connect the menu choice to your macro. The OnAction property
            ' should be set to the name of your macro.
            .OnAction = "HideAnnotations"
            .Tag = "HideAnnotations"
            .TooltipText = "Hide Annotations"
      End With
      Set NewControl = ToolsMenu("Tools").Controls.Add(Type:=msoControlButton, Before:=1)
      With NewControl
            .DescriptionText = "Shows Objects named as Annotation SHAPE XX."
            ' Name the command.
            .Caption = "Show Annotations"
            ' Connect the menu choice to your macro. The OnAction property
            ' should be set to the name of your macro.
            .OnAction = "ShowAnnotations"
            .Tag = "ShowAnnotations"
            .TooltipText = "Show Annotations"
      End With
End Sub

Sub Auto_Close()
      Dim oControl As Object
      Dim ToolsMenu As CommandBarControls
      ' Get an object reference to a command bar.
      Set ToolsMenu = Application.CommandBars("Tools").Controls
      ToolsMenu.Item("Hide Annotations").Delete
      ToolsMenu.Item("Show Annotations").Delete
End Sub

Sub HideAnnotations()
    ' Variable declarations.
    Dim oSl As Slide
    Dim oOldSlide, oNewSlide, oOldNotes, oSh  As Shape
    Dim mySlideNum As Integer
    Dim SourceSlides, x, lyly As Long
    Dim SourceView, answer As Integer
    Dim myFileName, myFileNameTmp, TestArray() As String
    Dim Response As String
    Dim myAnnotationName() As String
    Dim dlgOpen As FileDialog

    ' Check to see whether a presentation is open.
    If Presentations.Count <> 0 Then
        If ActiveWindow.ViewType <> ppViewNormal Then
            ActiveWindow.ViewType = ppViewNormal
        End If
    Else
        MsgBox "No presentation open. Open a presentation and " _
            & "run the macro again.", vbExclamation
    End If
    
    myFileName = Application.ActivePresentation.FullName
    TestArray() = Split(myFileName, ".")
    myFileNameTmp = TestArray(0)

    ' Stores the current view of the source presentation.
    SourceView = ActiveWindow.ViewType
    
    If ActivePresentation.Saved <> msoTrue Then
        Response = MsgBox("Your presentation is not saved. Would you like to save a copy?", vbYesNo)
        If Response = vbYes Then
            ActivePresentation.SaveCopyAs (myFileNameTmp & "_orig")
            ActivePresentation.Close
            Application.Presentations.Open (myFileName)
        End If
    End If
    
    ' Count the number of slides in source presentation.
    SourceSlides = ActivePresentation.Slides.Count

    For x = 1 To SourceSlides
        Set oSl = ActivePresentation.Slides(x)
        ' Hiding annotations
        For Each oSh In oSl.Shapes
                myAnnotationName() = Split(oSh.Name, " ", , vbTextCompare)
                If (myAnnotationName(0) = "Annotation") Then
                    oSh.Visible = msoFalse
                End If
        Next
    Next
    ActivePresentation.Save
End Sub

Sub ShowAnnotations()
    ' Variable declarations.
    Dim oSl As Slide
    Dim oOldSlide, oNewSlide, oOldNotes, oSh  As Shape
    Dim mySlideNum As Integer
    Dim SourceSlides, x, lyly As Long
    Dim SourceView, answer As Integer
    Dim myFileName, myFileNameTmp, TestArray() As String
    Dim Response As String
    Dim myAnnotationName() As String
    Dim dlgOpen As FileDialog

    ' Check to see whether a presentation is open.
    If Presentations.Count <> 0 Then
        If ActiveWindow.ViewType <> ppViewNormal Then
            ActiveWindow.ViewType = ppViewNormal
        End If
    Else
        MsgBox "No presentation open. Open a presentation and " _
            & "run the macro again.", vbExclamation
    End If
    
    myFileName = Application.ActivePresentation.FullName
    TestArray() = Split(myFileName, ".")
    myFileNameTmp = TestArray(0)

    ' Stores the current view of the source presentation.
    SourceView = ActiveWindow.ViewType
    
    If ActivePresentation.Saved <> msoTrue Then
        Response = MsgBox("Your presentation is not saved. Would you like to save a copy?", vbYesNo)
        If Response = vbYes Then
            ActivePresentation.SaveCopyAs (myFileNameTmp & "_orig")
            ActivePresentation.Close
            Application.Presentations.Open (myFileName)
        End If
    End If
    
    ' Count the number of slides in source presentation.
    SourceSlides = ActivePresentation.Slides.Count

    ' Loop through all the slides and copy them to destination one by one.
    For x = 1 To SourceSlides
        Set oSl = ActivePresentation.Slides(x)
        For Each oSh In oSl.Shapes
                myAnnotationName() = Split(oSh.Name, " ", , vbTextCompare)
                If (myAnnotationName(0) = "Annotation") Then
                    oSh.Visible = msoTrue
                    oSh.ZOrder (msoBringToFront)
                End If
            If oSh.Type = msoMedia Then
                oSh.Visible = msoTrue
            End If
            'Testing to see if all audio can be hidden
        Next
    Next
    ActivePresentation.Save
End Sub
