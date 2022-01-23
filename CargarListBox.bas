Attribute VB_Name = "CargarListBox"
Public Sub CargarlstCampoDisponible(NombreTabla As String)
    
    Dim Tabla As TableDef
    Dim Campo As Field
    Dim Indice As Index
    Dim i As Integer
    
    Set Tabla = dbBuilder.TableDefs(NombreTabla)
    For Each Campo In Tabla.Fields
        If Not Campo.Name = "CampoProvisorio" Then
            CargaIndice.lstCampoDisponible.AddItem Campo.Name
        End If
    Next
    
    i = 0
    While i < CargaIndice.lstCampoDisponible.ListCount
        CargaIndice.lstCampoDisponible.ListIndex = i
        For Each Indice In Tabla.Indexes
            If Indice.Primary = True Then
                Dim CampoIndezado As Field
                For Each CampoIndezado In Indice.Fields
                    If CampoIndezado.Name = CargaIndice.lstCampoDisponible.Text Then
                        Call PasarDatosListBox(CargaIndice.lstCampoDisponible, CargaIndice.lstCampoIndice)
                        If Not CargaIndice.lstCampoDisponible.ListCount = i Then
                            i = i - 1
                        End If
                        Exit For
                    End If
                Next CampoIndezado
            End If
        Next Indice
        i = i + 1
    Wend
    
End Sub

Public Sub VaciarListBox()

    With CargaIndice
        .lstCampoDisponible.Clear
        .lstCampoIndice.Clear
        '.cmdAgregar.Enabled = False
        '.cmdEliminar.Enabled = False
        '.cmdGuardar.Enabled = False
    End With

End Sub

Public Sub PasarDatosListBox(ListOrigen As ListBox, ListDestino As ListBox)

    If Not ListOrigen.Text = "" Then
        ListDestino.AddItem ListOrigen.Text
        ListOrigen.RemoveItem ListOrigen.ListIndex
    End If

End Sub
