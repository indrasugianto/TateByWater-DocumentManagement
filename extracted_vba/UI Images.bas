' Component: UI Images
' Type: module
' Lines: 11
' ============================================================

Option Compare Database

Public Sub GetallImages()
    Dim img As IPictureDisp
    Set img = Application.CommandBars.GetImageMso("FileSave", 16, 16)
    stdole.SavePicture img, CurrentProject.path & "/Savebtn.bmp"
    
    Set img = Application.CommandBars.GetImageMso("FileSave", 16, 16)
    stdole.SavePicture img, CurrentProject.path & "/Savebtn.bmp"
    
End Sub