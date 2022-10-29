Option Explicit

'The variable for holding the tables path & filename
Dim DiseaseTable_Filename As String
Dim CausalAgentTable_Filename As String
Dim DiseaseCausalTable_Filename As String
Dim SpeciesTable_Filename As String


'The  multidimensional arrays to hold all columns in each every table
Dim Tabel_diseases() As Variant
Dim Tabel_Causal() As Variant
Dim Tabel_Disease_Causal() As Variant
'
Dim Tabel_species_hierarchy() As Variant
Dim Tabel_species_group() As Variant
Dim Tabel_Aquatic_species_order() As Variant
Dim Tabel_Terristrial_species_order() As Variant

'
Dim ListBox_Item_no As Integer


'This function takes the disease id and return a list of relevant causal agents
Function get_causal_agents(Disease_id As Integer) As Variant
    '
    Dim this_Disease_Causal_Agents() As Variant
    Dim Lst_Causal_name() As Variant
    Dim i, k, arrSize As Integer
    '
    arrSize = 0
    '
    For i = 1 To UBound(Tabel_Disease_Causal(0))
        If Int(Tabel_Disease_Causal(0)(i, 1)) = Disease_id Then
            arrSize = arrSize + 1
            ReDim Preserve this_Disease_Causal_Agents(arrSize)
            this_Disease_Causal_Agents(arrSize - 1) = Int(Tabel_Disease_Causal(1)(i, 1))
        End If
    Next i
    '
    If arrSize = 0 Then
        ReDim Preserve Lst_Causal_name(arrSize + 1)
        Lst_Causal_name(arrSize) = ">>>>>>> No Causal Agent <<<<<<"
    Else
        ReDim Preserve this_Disease_Causal_Agents(arrSize - 1)
        arrSize = 0
        '
        For i = 0 To UBound(this_Disease_Causal_Agents)
            For k = 1 To UBound(Tabel_Causal(0))
                If Int(Tabel_Causal(0)(k, 1)) = this_Disease_Causal_Agents(i) Then
                    arrSize = arrSize + 1
                    ReDim Preserve Lst_Causal_name(arrSize)
                    Lst_Causal_name(arrSize - 1) = Tabel_Causal(1)(k, 1)
                End If
            Next k
        Next i
    End If
    get_causal_agents = Lst_Causal_name
End Function
'
Function OpenExelFiles(fileNameStr As String) As String
    ' Create and set the file dialog object.
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear      ' Clear all the filters (if applied before).
        
        ' Give the dialog box a title, word for doc or Excel for excel files.
        .Title = fileNameStr
        
        ' Apply filter to show only a particular type of files.
        ' For example, *.doc? to show only word files or
        ' *.xlsx? to show only excel files.
        .Filters.Add "Exel Files", "*.xlsx?", 1
        
        ' Do not allow users to select more than one file.
        ' Set the value as "True" to select multiple files.
        .AllowMultiSelect = False
    
        ' Show the file.
        If .Show = True Then
            Debug.Print .SelectedItems(1)           ' Get the complete file path.
            Debug.Print Dir(.SelectedItems(1))      ' Get the file name.
            OpenExelFiles = .SelectedItems(1)
        Else
            OpenExelFiles = ""
        End If
    End With
    
End Function
'
Private Sub CommandButton_get_tables_Click()
    Call Get_Disease_Causal
    Call Get_Species
    '
    With UserForm_DiseaseCausal
        .ListBox_Diseases.List = Tabel_diseases(1)
        .Label15 = Str(ListBox_Item_no)
        .CommandButton_openDisease.Enabled = False
        .CommandButton_Open_Disease_Causal.Enabled = False
        .CommandButton_open_species.Enabled = False
        .CommandButton_opencausal.Enabled = False
    End With
   
End Sub

'This procedure open each table file and extract all relevant columns and save them in a multidimensional array
Private Sub Get_Disease_Causal()
    Dim wkbook_disease As Workbook
    Dim wkbook_dcausal As Workbook
    Dim wkbook_disease_causal As Workbook
    '
    Dim sheet_disease As Worksheet
    Dim sheet_causal As Worksheet
    Dim sheet_disease_causal As Worksheet
    '
    'Disease Table Lists
    Dim Lst_disease_id() As Variant
    Dim Lst_disease_name() As Variant
    Dim Lst_disease_group() As Variant
    Dim Lst_disease_OIE_Listed() As Variant
    Dim Lst_disease_Non_OIE_listed() As Variant
    Dim Lst_disease_Emerging() As Variant
    Dim Lst_disease_Self_declaration() As Variant
    Dim Lst_disease_Concern_aquatic() As Variant
    Dim Diseses_sum As Integer
    
    'Causal Agent Lists
    Dim Lst_causal_id() As Variant
    Dim Lst_Causal_name() As Variant
    Dim Lst_causal_agnt_type() As Variant
    
    'Disease Causal Agent Lists
    Dim Lst_Diseasecausal_Disease_id() As Variant
    Dim Lst_Diseasecausal_Causal_id() As Variant
    
    ' Get the workbooks of all table files
    Application.ScreenUpdating = False
    Set wkbook_disease = Workbooks.Open(DiseaseTable_Filename)
    '
    Set wkbook_dcausal = Workbooks.Open(CausalAgentTable_Filename)
    '
    Set wkbook_disease_causal = Workbooks.Open(DiseaseCausalTable_Filename)
    '
    ActiveWindow.Visible = False
    Application.ScreenUpdating = True
       
    ' Get sheets of each workbook
    Set sheet_disease = wkbook_disease.Sheets(1)
    Set sheet_causal = wkbook_dcausal.Sheets(1)
    Set sheet_disease_causal = wkbook_disease_causal.Sheets(1)
    
    ' get all relevant columns in Disease table
    With sheet_disease
        Lst_disease_id = .Range("a2", .Range("a2").End(xlDown)).Value                             'Unique_id          Tabel_diseases(0)(i,1)
        Lst_disease_name = .Range("d2", .Range("d2").End(xlDown)).Value                           'Disease Name       Tabel_diseases(1)(i,1)
        Diseses_sum = Application.CountA(Lst_disease_name)
        '
        Lst_disease_group = .Range(.Cells(2, 9), .Cells(Diseses_sum + 1, 9)).Value2               'Disease Group      Tabel_diseases(2)(i,1)
        Lst_disease_OIE_Listed = .Range(.Cells(2, 10), .Cells(Diseses_sum + 1, 10)).Value2        'OIE-Listed disease Tabel_diseases(3)(i,1)
        Lst_disease_Non_OIE_listed = .Range(.Cells(2, 11), .Cells(Diseses_sum + 1, 11)).Value2    'Non-OIE_listed     Tabel_diseases(4)(i,1)
        Lst_disease_Emerging = .Range(.Cells(2, 12), .Cells(Diseses_sum + 1, 12)).Value2          'Emerging disease   Tabel_diseases(5)(i,1)
        Lst_disease_Self_declaration = .Range(.Cells(2, 15), .Cells(Diseses_sum + 1, 15)).Value2  'Self-declaration   Tabel_diseases(6)(i,1)
        Lst_disease_Concern_aquatic = .Range(.Cells(2, 17), .Cells(Diseses_sum + 1, 17)).Value2   'Concern_aquatic    Tabel_diseases(7)(i,1)
    
    End With
    
    ' Get all relevant columns in causal table
    Lst_causal_id = sheet_causal.Range("a2", sheet_causal.Range("a2").End(xlDown)).Value          'Unique_id         Tabel_diseases(0)(i,1)
    Lst_Causal_name = sheet_causal.Range("b2", sheet_causal.Range("b2").End(xlDown)).Value        'Agent Name        Tabel_diseases(1)(i,1)
    Lst_causal_agnt_type = sheet_causal.Range("c2", sheet_causal.Range("c2").End(xlDown)).Value   'Agent Type        Tabel_diseases(2)(i,1)
    
    ' Get all relevant columns in Disease & Causal table
    Lst_Diseasecausal_Disease_id = sheet_disease_causal.Range("b2", sheet_disease_causal.Range("b2").End(xlDown)).Value       'Disease Unique_id          Tabel_diseases(0)(i,1)
    Lst_Diseasecausal_Causal_id = sheet_disease_causal.Range("a2", sheet_disease_causal.Range("a2").End(xlDown)).Value        'Causal Agent Unique_id     Tabel_diseases(1)(i,1)
    
    'Add all columns in disease table to a multidimensional array
    Tabel_diseases() = Array(Lst_disease_id, _
                                Lst_disease_name, _
                                Lst_disease_group, _
                                Lst_disease_OIE_Listed, _
                                Lst_disease_Non_OIE_listed, _
                                Lst_disease_Emerging, _
                                Lst_disease_Self_declaration, _
                                Lst_disease_Concern_aquatic)
                                
    'Add all columns in causal agents table to a multidimensional array
    Tabel_Causal() = Array(Lst_causal_id, _
                            Lst_Causal_name, _
                            Lst_causal_agnt_type)
                            
    'Add all columns in disease & causal agents table to a multidimensional array
    Tabel_Disease_Causal() = Array(Lst_Diseasecausal_Disease_id, _
                                    Lst_Diseasecausal_Causal_id)
       
    wkbook_disease.Close SaveChanges:=False
    wkbook_dcausal.Close SaveChanges:=False
    wkbook_disease_causal.Close SaveChanges:=False
    
End Sub

Private Sub CommandButton_open_species_Click()
    SpeciesTable_Filename = OpenExelFiles("Open Species table")
    UserForm_DiseaseCausal.Label_species.Caption = SpeciesTable_Filename
End Sub

Private Sub CommandButton_openDisease_Click()
    DiseaseTable_Filename = OpenExelFiles("Open Diseases table")
    UserForm_DiseaseCausal.Label_disease.Caption = DiseaseTable_Filename
End Sub

Private Sub CommandButton_opencausal_Click()
    CausalAgentTable_Filename = OpenExelFiles("Open Causal Agents table")
    UserForm_DiseaseCausal.Label_causal.Caption = CausalAgentTable_Filename
End Sub

Private Sub CommandButton_Open_Disease_Causal_Click()
    DiseaseCausalTable_Filename = OpenExelFiles("Open Diseases&Causal table")
    UserForm_DiseaseCausal.Label_diseae_causal.Caption = DiseaseCausalTable_Filename
End Sub
'
Private Sub ListBox_Diseases_Click()
    '
    With UserForm_DiseaseCausal
        ListBox_Item_no = .ListBox_Diseases.ListIndex + 1
        .Label_disease_group = Tabel_diseases(2)(ListBox_Item_no, 1)
        .Label_OIE_Listed = Tabel_diseases(3)(ListBox_Item_no, 1)
        .Label_Non_OIE_listed = Tabel_diseases(4)(ListBox_Item_no, 1)
        .Label_Emerging_Disease = Tabel_diseases(5)(ListBox_Item_no, 1)
        .Label_Self_Declaration = Tabel_diseases(6)(ListBox_Item_no, 1)
        .Label_is_aquatic = Tabel_diseases(7)(ListBox_Item_no, 1)
        .Label15 = Str(UBound(Tabel_diseases(0))) & "/" & Str(ListBox_Item_no)
        '
        .ListBox_Causals.List = get_causal_agents(Int(Tabel_diseases(0)(ListBox_Item_no, 1)))
        '
        If Tabel_diseases(7)(ListBox_Item_no, 1) = 1 Then
            UserForm_DiseaseCausal.ListBox_species.List = Tabel_Aquatic_species_order(1)
        Else
            UserForm_DiseaseCausal.ListBox_species.List = Tabel_Terristrial_species_order(1)
        End If
    End With
    
End Sub

'
Private Sub Get_Species()
    Dim wkbook_species_hierarchy As Workbook
    Dim wkbook_species_group As Workbook
    '
    Dim sheet_species As Worksheet
    Dim sheet_species_lookup_sheet As Worksheet
    
    ' Get the workbooks of all table files and the sheet
    Application.ScreenUpdating = False
    '
    Set wkbook_species_group = Workbooks.Open(SpeciesTable_Filename)
    '
    ActiveWindow.Visible = False
    Application.ScreenUpdating = True
    '
    Set sheet_species = wkbook_species_group.Sheets(1)

    ' get all relevant columns in Disease table
    With sheet_species
        
        'Rest Filter
        .AutoFilterMode = False
        
        'Reset filter on terrestrial animal
        .Range("$A$1:$U$3580").AutoFilter Field:=6, Criteria1:=Array( _
                "amp", "amp (fau)", "api", "avi", "bov", "buf", "can", "cap", "cer", "cml", "cru", _
                    "cru (fau)", "equ", "fel", "lep", "mol", "mol (fau)", "o/c", "ovi", "pis", "pis (fau)", _
                        "sui"), Operator:=xlFilterValues
      
        'Add all columns in disease table to a multidimensional array
        Tabel_Terristrial_species_order() = Array( _
                                .Range("c2", .Range("c2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("d2", .Range("d2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("f2", .Range("f2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("h2", .Range("h2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("n2", .Range("n2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value)
        'Rest Filter
        .AutoFilterMode = False
         
        'Reset filter on aquatic animal
        .Range("$A$1:$U$3580").AutoFilter Field:=6, Criteria1:=Array( _
                "amp", "amp (fau)", "cru", "cru (fau)", "mol", "mol (fau)", _
                 "pis", "pis (fau)"), Operator:=xlFilterValues
       
         Tabel_Aquatic_species_order() = Array( _
                                .Range("c2", .Range("c2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("d2", .Range("d2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("f2", .Range("f2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("h2", .Range("h2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value, _
                                .Range("n2", .Range("n2").End(xlDown)).SpecialCells(xlCellTypeVisible).Value)
        'Rest Filter
        .AutoFilterMode = False

    End With
    '
    wkbook_species_group.Close SaveChanges:=False

    
End Sub


Private Sub UserForm_Click()

End Sub
