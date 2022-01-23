Attribute VB_Name = "CargarComboBox"
Public Sub CargarcmbNombreTabla(FormularioOrigen As Form)

    Dim Tabla As TableDef
    For Each Tabla In dbBuilder.TableDefs
        If Not Tabla.Name Like "MSys*" Then
            FormularioOrigen.cmbNombreTabla.AddItem Tabla.Name
        End If
    Next

End Sub

Public Sub CargarcmbTipoCampo()

    With CargaCampo.cmbTipo
        .AddItem "Texto", 0
        .AddItem "Moneda", 1
        .AddItem "Long", 2
        .AddItem "Integer", 3
        .AddItem "Byte", 4
        .AddItem "Date/Time", 5
        .AddItem "Boleano", 6
        .AddItem "Single", 7
        .AddItem "Double", 8
    End With

End Sub

Public Sub CargarcmbTablaOrigenyDestino()

    Dim Tabla As TableDef
    For Each Tabla In dbBuilder.TableDefs
        If Not Tabla.Name Like "MSys*" Then
            CargaRelacion.cmbTablaOrigen.AddItem Tabla.Name
            CargaRelacion.cmbTablaDestino.AddItem Tabla.Name
        End If
    Next

End Sub

Public Sub CargarcmbCampoOrigen(TablaOrigen As String)

    Dim Campo As Field
    Dim Tabla As TableDef
    
    Set Tabla = dbBuilder.TableDefs(TablaOrigen)
    For Each Campo In Tabla.Fields
        If Not Campo.Name = "CampoProvisorio" Then
            CargaRelacion.cmbCampoOrigen.AddItem Campo.Name
        End If
    Next

End Sub

Public Sub CargarcmbCampoDestino(TablaDestino As String)

    Dim Campo As Field
    Dim Tabla As TableDef
    
    Set Tabla = dbBuilder.TableDefs(TablaDestino)
    For Each Campo In Tabla.Fields
        If Not Campo.Name = "CampoProvisorio" Then
            CargaRelacion.cmbCampoDestino.AddItem Campo.Name
        End If
    Next

End Sub

Public Sub CargarcmbOrdenCampo(NombreTabla As String)

    Dim Campo As Field
    Dim Tabla As TableDef
    
    Set Tabla = dbBuilder.TableDefs(NombreTabla)
    For Each Campo In Tabla.Fields
        If Not Campo.Name = "CampoProvisorio" Then
            ListadoRegistros.cmbOrdenCampo.AddItem Campo.Name
        End If
    Next

End Sub

Public Sub CargarcmbTablaImportar()

    Dim Tabla As TableDef
    For Each Tabla In dbBuilder.TableDefs
        If Not Tabla.Name Like "MSys*" Then
            ImportarXLS.cmbTabla.AddItem Tabla.Name
        End If
    Next

End Sub
