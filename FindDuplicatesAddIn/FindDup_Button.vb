Public Class FindDup_Button
  Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Private oForm As frmFindDups

  Public Sub New()

  End Sub

  Protected Overrides Sub OnClick()
        'fire up the form
        oForm = New frmFindDups(My.ArcMap.Application.Document)
        oForm.TopMost = True
        oForm.Show()
    My.ArcMap.Application.CurrentTool = Nothing
  End Sub

  Protected Overrides Sub OnUpdate()
        Enabled = My.ArcMap.Application IsNot Nothing
    End Sub
    
End Class
