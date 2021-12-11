
Public Sub init(confManager As vtkConfigurationManager, confNumber As Integer)
    Set m_parent = confManager
    m_number = confNumber
End Sub

Public Sub addReferencesToWorkbook(Wb As Workbook)
    Dim ref As vtkReference
    For Each ref In Me.references
        ref.addToWorkbook Wb
    Next
End Sub


