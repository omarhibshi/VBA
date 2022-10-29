Option Explicit

'The variable for holding the tables path & filename
Dim DiseaseTable_Filename As String
Dim CausalAgentTable_Filename As String
Dim DiseaseCausalTable_Filename As String
Dim SpeciesTable_Filename As String
Dim SpeciesHierarchiesTable_Filename As String
Dim DiseaseSpeciesTable_Filename As String

'The  multidimensional arrays to hold all columns in each every table
Dim Tabel_diseases_terre() As Variant
Dim Tabel_diseases_aqua() As Variant
'
Dim Tabel_Causal() As Variant
Dim Tabel_Disease_Causal() As Variant
'
Dim Tabel_species_hierarchy() As Variant
Dim Tabel_species_group_Aquatic() As Variant
Dim Tabel_species_group_Terristrial() As Variant
'

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
    If arrSize > 0 Then
        ReDim Preserve Lst_Causal_name(arrSize - 1)
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
Private Sub ComboBox_disease_Change()
    If UserForm_DiseaseCausal.ComboBox_disease.ListIndex > -1 Then
        UserForm_DiseaseCausal.ListBox_Diseases.ListIndex = UserForm_DiseaseCausal.ComboBox_disease.ListIndex
    End If
    Call ListBox_Diseases_Click
End Sub


'
Private Sub CommandButton_get_tables_Click()
    Call Get_Disease_Causal
    Call Get_Species_Group
    '
    With UserForm_DiseaseCausal
        .CommandButton_openDisease.Enabled = False
        .CommandButton_opencausal.Enabled = False
        .CommandButton_Open_Disease_Causal.Enabled = False
        .CommandButton_open_species.Enabled = False
        .CommandButton_species_hr.Enabled = False
        .CommandButton_disease_species.Enabled = False
        .OptionButton_Aqaua.Enabled = False
        .OptionButton_Terre.Enabled = False
        .Label1_animalType.Enabled = False
        '
        If .OptionButton_Terre Then
            .ListBox_Diseases.List = Tabel_diseases_terre(1)
            .ComboBox_disease.List = Tabel_diseases_terre(1)
        Else
            .ListBox_Diseases.List = Tabel_diseases_aqua(1)
            .ComboBox_disease.List = Tabel_diseases_aqua(1)
        End If
    End With
End Sub

'This procedure open each table file and extract all relevant columns and save them in a multidimensional array
Private Sub Get_Disease_Causal()
    '
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
    '
    Dim Row_count As Long
    Dim i, arraySize As Integer
    
    'Causal Agent Lists
    Dim Lst_causal_id() As Variant
    Dim Lst_Causal_name() As Variant
    Dim Lst_causal_agnt_type() As Variant
    
    'Disease Causal Agent Lists
    Dim Lst_Diseasecausal_Disease_id() As Variant
    Dim Lst_Diseasecausal_Causal_id() As Variant
    '
    
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
    
        'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("a2", .Range("a2").End(xlDown)).Count
        
        'Terristrial disease list
        arraySize = 0
        
        For i = 2 To Row_count - 1
            If .Cells(i, 17).Value = 0 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_disease_id(arraySize)
                ReDim Preserve Lst_disease_name(arraySize)
                ReDim Preserve Lst_disease_group(arraySize)
                ReDim Preserve Lst_disease_OIE_Listed(arraySize)
                ReDim Preserve Lst_disease_Non_OIE_listed(arraySize)
                ReDim Preserve Lst_disease_Emerging(arraySize)
                ReDim Preserve Lst_disease_Self_declaration(arraySize)
                ReDim Preserve Lst_disease_Concern_aquatic(arraySize)
                '
                Lst_disease_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_disease_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_disease_group(arraySize - 1) = .Cells(i, 9).Value
                Lst_disease_OIE_Listed(arraySize - 1) = .Cells(i, 10).Value
                Lst_disease_Non_OIE_listed(arraySize - 1) = .Cells(i, 11).Value
                Lst_disease_Emerging(arraySize - 1) = .Cells(i, 12).Value
                Lst_disease_Self_declaration(arraySize - 1) = .Cells(i, 15).Value
                Lst_disease_Concern_aquatic(arraySize - 1) = .Cells(i, 17).Value

           End If
        Next i
        ReDim Preserve Lst_disease_id(arraySize - 1)
        ReDim Preserve Lst_disease_name(arraySize - 1)
        ReDim Preserve Lst_disease_group(arraySize - 1)
        ReDim Preserve Lst_disease_OIE_Listed(arraySize - 1)
        ReDim Preserve Lst_disease_Non_OIE_listed(arraySize - 1)
        ReDim Preserve Lst_disease_Emerging(arraySize - 1)
        ReDim Preserve Lst_disease_Self_declaration(arraySize - 1)
        ReDim Preserve Lst_disease_Concern_aquatic(arraySize - 1)
        
        Tabel_diseases_terre() = Array( _
                        Lst_disease_id, _
                        Lst_disease_name, _
                        Lst_disease_group, _
                        Lst_disease_OIE_Listed, _
                        Lst_disease_Non_OIE_listed, _
                        Lst_disease_Emerging, _
                        Lst_disease_Self_declaration, _
                        Lst_disease_Concern_aquatic)
                        
        Erase Lst_disease_id
        Erase Lst_disease_name
        Erase Lst_disease_group
        Erase Lst_disease_OIE_Listed
        Erase Lst_disease_Non_OIE_listed
        Erase Lst_disease_Emerging
        Erase Lst_disease_Self_declaration
        Erase Lst_disease_Concern_aquatic
                        
        'Aqua disease list
        arraySize = 0
                        
        For i = 2 To Row_count - 1
            If .Cells(i, 17).Value = 1 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_disease_id(arraySize)
                ReDim Preserve Lst_disease_name(arraySize)
                ReDim Preserve Lst_disease_group(arraySize)
                ReDim Preserve Lst_disease_OIE_Listed(arraySize)
                ReDim Preserve Lst_disease_Non_OIE_listed(arraySize)
                ReDim Preserve Lst_disease_Emerging(arraySize)
                ReDim Preserve Lst_disease_Self_declaration(arraySize)
                ReDim Preserve Lst_disease_Concern_aquatic(arraySize)
                '
                Lst_disease_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_disease_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_disease_group(arraySize - 1) = .Cells(i, 9).Value
                Lst_disease_OIE_Listed(arraySize - 1) = .Cells(i, 10).Value
                Lst_disease_Non_OIE_listed(arraySize - 1) = .Cells(i, 11).Value
                Lst_disease_Emerging(arraySize - 1) = .Cells(i, 12).Value
                Lst_disease_Self_declaration(arraySize - 1) = .Cells(i, 15).Value
                Lst_disease_Concern_aquatic(arraySize - 1) = .Cells(i, 17).Value
            End If
        Next i
        '
        ReDim Preserve Lst_disease_id(arraySize - 1)
        ReDim Preserve Lst_disease_name(arraySize - 1)
        ReDim Preserve Lst_disease_group(arraySize - 1)
        ReDim Preserve Lst_disease_OIE_Listed(arraySize - 1)
        ReDim Preserve Lst_disease_Non_OIE_listed(arraySize - 1)
        ReDim Preserve Lst_disease_Emerging(arraySize - 1)
        ReDim Preserve Lst_disease_Self_declaration(arraySize - 1)
        ReDim Preserve Lst_disease_Concern_aquatic(arraySize - 1)
        
        Tabel_diseases_aqua() = Array( _
                        Lst_disease_id, _
                        Lst_disease_name, _
                        Lst_disease_group, _
                        Lst_disease_OIE_Listed, _
                        Lst_disease_Non_OIE_listed, _
                        Lst_disease_Emerging, _
                        Lst_disease_Self_declaration, _
                        Lst_disease_Concern_aquatic)
     
    End With
    
    ' Get all relevant columns in causal table
    Lst_causal_id = sheet_causal.Range("a2", sheet_causal.Range("a2").End(xlDown)).Value          'Unique_id         Tabel_diseases(0)(i,1)
    Lst_Causal_name = sheet_causal.Range("b2", sheet_causal.Range("b2").End(xlDown)).Value        'Agent Name        Tabel_diseases(1)(i,1)
    Lst_causal_agnt_type = sheet_causal.Range("c2", sheet_causal.Range("c2").End(xlDown)).Value   'Agent Type        Tabel_diseases(2)(i,1)
    
    ' Get all relevant columns in Disease & Causal table
    Lst_Diseasecausal_Disease_id = sheet_disease_causal.Range("b2", sheet_disease_causal.Range("b2").End(xlDown)).Value       'Disease Unique_id          Tabel_diseases(0)(i,1)
    Lst_Diseasecausal_Causal_id = sheet_disease_causal.Range("a2", sheet_disease_causal.Range("a2").End(xlDown)).Value        'Causal Agent Unique_id     Tabel_diseases(1)(i,1)
    
                                
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
Private Sub CommandButton_openDisease_Click()
    DiseaseTable_Filename = OpenExelFiles("Open Diseases table")
    UserForm_DiseaseCausal.Label_disease.Caption = DiseaseTable_Filename
End Sub
Private Sub CommandButton_opencausal_Click()
    CausalAgentTable_Filename = OpenExelFiles("Open Causal Agents table")
    UserForm_DiseaseCausal.Label_causal.Caption = CausalAgentTable_Filename
End Sub

Private Sub CommandButton_Open_Disease_Causal_Click()
    DiseaseCausalTable_Filename = OpenExelFiles("Open Diseases and Causal Agents table")
    UserForm_DiseaseCausal.Label_diseae_causal.Caption = DiseaseCausalTable_Filename
End Sub
Private Sub CommandButton_species_hr_Click()
    SpeciesHierarchiesTable_Filename = OpenExelFiles("Open Species hierarchies table")
    UserForm_DiseaseCausal.Label_SpeciesHierarchies.Caption = SpeciesHierarchiesTable_Filename
End Sub
Private Sub CommandButton_open_species_Click()
    SpeciesTable_Filename = OpenExelFiles("Open Species And Groups table")
    UserForm_DiseaseCausal.Label_species.Caption = SpeciesTable_Filename
End Sub
Private Sub CommandButton_disease_species_Click()
    DiseaseSpeciesTable_Filename = OpenExelFiles("Open Diseases affects Specieses table")
    UserForm_DiseaseCausal.Label1_DisesesSpecies.Caption = DiseaseSpeciesTable_Filename
End Sub
Private Sub ListBox_Causals_Click()
    UserForm_DiseaseCausal.TextBox_causal_name.Value = UserForm_DiseaseCausal.ListBox_Causals.Value
End Sub
'
Private Sub ListBox_Diseases_Click()
    '
    Dim ListBox_Item_no As Integer
    
    With UserForm_DiseaseCausal
        ListBox_Item_no = .ListBox_Diseases.ListIndex
        If .OptionButton_Terre Then
            .Label_disease_group = Tabel_diseases_terre(2)(ListBox_Item_no)
            .Label_OIE_Listed = Tabel_diseases_terre(3)(ListBox_Item_no)
            .Label_Non_OIE_listed = Tabel_diseases_terre(4)(ListBox_Item_no)
            .Label_Emerging_Disease = Tabel_diseases_terre(5)(ListBox_Item_no)
            .Label_Self_Declaration = Tabel_diseases_terre(6)(ListBox_Item_no)
            .Label_is_aquatic = Tabel_diseases_terre(7)(ListBox_Item_no)
            .Label15 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_diseases_terre(0)))
            '
            .ListBox_Causals.List = get_causal_agents(Int(Tabel_diseases_terre(0)(ListBox_Item_no)))
            .TextBox_causal_name.Value = ""
            '
            If Tabel_diseases_terre(7)(ListBox_Item_no) = 1 Then
                UserForm_DiseaseCausal.ListBox_species.List = Tabel_species_group_Aquatic(1)
            Else
                UserForm_DiseaseCausal.ListBox_species.List = Tabel_species_group_Terristrial(1)
            End If
        Else
            .Label_disease_group = Tabel_diseases_aqua(2)(ListBox_Item_no)
            .Label_OIE_Listed = Tabel_diseases_aqua(3)(ListBox_Item_no)
            .Label_Non_OIE_listed = Tabel_diseases_aqua(4)(ListBox_Item_no)
            .Label_Emerging_Disease = Tabel_diseases_aqua(5)(ListBox_Item_no)
            .Label_Self_Declaration = Tabel_diseases_aqua(6)(ListBox_Item_no)
            .Label_is_aquatic = Tabel_diseases_aqua(7)(ListBox_Item_no)
            .Label15 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_diseases_aqua(0)))
            '
            .ListBox_Causals.List = get_causal_agents(Int(Tabel_diseases_aqua(0)(ListBox_Item_no)))
            .TextBox_causal_name.Value = ""
            '
            If Tabel_diseases_aqua(7)(ListBox_Item_no) = 1 Then
                UserForm_DiseaseCausal.ListBox_species.List = Tabel_species_group_Aquatic(1)
            Else
                UserForm_DiseaseCausal.ListBox_species.List = Tabel_species_group_Terristrial(1)
            End If
        End If
    End With
End Sub
'
Private Sub Get_Species_Group()
    '
    Dim wkbook_species_group As Workbook
    Dim sheet_species_group As Worksheet
    '
    'Disease Table Lists
    Dim Lst_species_group_id() As Variant
    Dim Lst_species_group_name() As Variant
    Dim Lst_species_group_parent_id() As Variant
    Dim Lst_species_group_hierarchy_id() As Variant
    Dim Lst_species_group_enTransl() As Variant
    '
    Dim Row_count As Long
    Dim i, arraySize As Integer
    
    ' Get the workbooks and sheets of all table files and the sheet
    Application.ScreenUpdating = False
    Set wkbook_species_group = Workbooks.Open(SpeciesTable_Filename)
    Set sheet_species_group = wkbook_species_group.Sheets(1)
    ActiveWindow.Visible = False
    Application.ScreenUpdating = True
    '
    With sheet_species_group
        
        'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("h2", .Range("h2").End(xlDown)).Count
        
        'Terristrial disease list
        arraySize = 0
        
        For i = 2 To Row_count - 1
            If InStr("34567", Trim(Str(.Cells(i, 8).Value))) > 0 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_species_group_id(arraySize)
                ReDim Preserve Lst_species_group_name(arraySize)
                ReDim Preserve Lst_species_group_parent_id(arraySize)
                ReDim Preserve Lst_species_group_hierarchy_id(arraySize)
                ReDim Preserve Lst_species_group_enTransl(arraySize)
                '
                Lst_species_group_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_species_group_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_species_group_parent_id(arraySize - 1) = .Cells(i, 9).Value
                Lst_species_group_hierarchy_id(arraySize - 1) = .Cells(i, 10).Value
                Lst_species_group_enTransl(arraySize - 1) = .Cells(i, 11).Value
           End If
        Next i
        ReDim Preserve Lst_species_group_id(arraySize - 1)
        ReDim Preserve Lst_species_group_name(arraySize - 1)
        ReDim Preserve Lst_species_group_parent_id(arraySize - 1)
        ReDim Preserve Lst_species_group_hierarchy_id(arraySize - 1)
        ReDim Preserve Lst_species_group_enTransl(arraySize - 1)
        
        Tabel_species_group_Terristrial() = Array( _
                        Lst_species_group_id, _
                        Lst_species_group_name, _
                        Lst_species_group_parent_id, _
                        Lst_species_group_hierarchy_id, _
                        Lst_species_group_enTransl)
                        
        Erase Lst_species_group_id
        Erase Lst_species_group_name
        Erase Lst_species_group_parent_id
        Erase Lst_species_group_hierarchy_id
        Erase Lst_species_group_enTransl
                        
        'Aqua disease list
        arraySize = 0
        
        'Add all columns in disease table to a multidimensional array
        For i = 2 To Row_count - 1
            If InStr("34567", Trim(Str(.Cells(i, 8).Value))) > 0 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_species_group_id(arraySize)
                ReDim Preserve Lst_species_group_name(arraySize)
                ReDim Preserve Lst_species_group_parent_id(arraySize)
                ReDim Preserve Lst_species_group_hierarchy_id(arraySize)
                ReDim Preserve Lst_species_group_enTransl(arraySize)
                '
                Lst_species_group_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_species_group_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_species_group_parent_id(arraySize - 1) = .Cells(i, 9).Value
                Lst_species_group_hierarchy_id(arraySize - 1) = .Cells(i, 10).Value
                Lst_species_group_enTransl(arraySize - 1) = .Cells(i, 11).Value
           End If
        Next i
        ReDim Preserve Lst_species_group_id(arraySize - 1)
        ReDim Preserve Lst_species_group_name(arraySize - 1)
        ReDim Preserve Lst_species_group_parent_id(arraySize - 1)
        ReDim Preserve Lst_species_group_hierarchy_id(arraySize - 1)
        ReDim Preserve Lst_species_group_enTransl(arraySize - 1)
        
        Tabel_species_group_Aquatic() = Array( _
                        Lst_species_group_id, _
                        Lst_species_group_name, _
                        Lst_species_group_parent_id, _
                        Lst_species_group_hierarchy_id, _
                        Lst_species_group_enTransl)
                        
        Erase Lst_species_group_id
        Erase Lst_species_group_name
        Erase Lst_species_group_parent_id
        Erase Lst_species_group_hierarchy_id
        Erase Lst_species_group_enTransl
        
    End With
    '
    wkbook_species_group.Close SaveChanges:=False

    
End Sub

Private Sub ListBox_species_Click()
    '
    Dim ListBox_Item_no As Integer
    
    With UserForm_DiseaseCausal
        ListBox_Item_no = .ListBox_species.ListIndex
        If .OptionButton_Terre Then
            .Label16 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_species_group_Terristrial(0)))
            '
        Else
            .Label16 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_species_group_Aquatic(0)))
        End If
    End With
End Sub

Private Sub OptionButton_Aqaua_Click()
    UserForm_DiseaseCausal.CommandButton_get_tables.Enabled = True
End Sub

Private Sub OptionButton_Terre_Click()
    UserForm_DiseaseCausal.CommandButton_get_tables.Enabled = True
End Sub

Private Sub UserForm_Click()

End Sub


