Attribute VB_Name = "ReferenceTable"
Option Explicit

'The variable for holding the tables path & filename
Public DiseaseTable_Filename As String
Public CausalAgentTable_Filename As String
Public DiseaseCausalTable_Filename As String
Public SpeciesTable_Filename As String
Public SpeciesHierarchiesTable_Filename As String
Public DiseaseSpeciesTable_Filename As String
Public SourceofInfectionTable_Filename As String
Public ControlMeasures_Filename As String
Public ControlMeasures_affect_species_Filename As String

'The  multidimensional arrays to hold all columns in each every table
Public Tabel_diseases_terre() As Variant
Public Tabel_diseases_aqua() As Variant
Public Tabel_diseases() As Variant
'
Public Tabel_Causal() As Variant
Public Tabel_Disease_Causal() As Variant
'
Public Tabel_ssource_ofInfection() As Variant
Public Tabel_Control_Measures() As Variant
Public Tabel_Control_Measures_domestic() As Variant
Public Tabel_Control_Measures_wild() As Variant
Public Tabel_Control_Measures_affect_species() As Variant
'
Public Tabel_species_hierarchy() As Variant
Public Tabel_species_group_Aquatic() As Variant
Public Tabel_species_group_Terristrial() As Variant
Public Tabel_species_groups() As Variant
Public Tabel_disease_affect_Species() As Variant
Public Tabel_susceptible_Species() As Variant


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

Function is_in_table(CM_id As Integer, lst As Variant) As Boolean
    Dim i As Integer
    '
    is_in_table = False
    For i = 0 To UBound(lst)
        If Int(lst(i)) = CM_id Then
            is_in_table = True
            Exit For
        End If
    Next i
End Function
Private Sub ComboBox_disease_Change()
    If UserForm_DiseaseCausal.ComboBox_disease.ListIndex > -1 Then
        UserForm_DiseaseCausal.ListBox_Diseases.ListIndex = UserForm_DiseaseCausal.ComboBox_disease.ListIndex
    End If
    Call ListBox_Diseases_Click
End Sub

'Get all reference table
Sub get_Reference_tables()

    DiseaseTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\diseases_20200504.xlsx"
    CausalAgentTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\causal_agent_20200323.xlsx"
    DiseaseCausalTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\disease_causal_agents_20191112.xlsx"
    SpeciesTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\species_and_groups_20200505.xlsx"
    SpeciesHierarchiesTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\species_hierarchies_types_20200403.xlsx"
    DiseaseSpeciesTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\disease_affect_species_20200427.xlsx"
    SourceofInfectionTable_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\source_of_infections_20190530.xlsx"
    ControlMeasures_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\control_measures_20190603.xlsx"
    ControlMeasures_affect_species_Filename = "C:\Users\asus\Documents\Oie\REFERENCE TABLES\control_measure_concerns_species_20191202.xlsx"
   
    Application.ScreenUpdating = False
    '
    Call Get_Disease_Causal
    'Call Get_Species_Group
    'Call Get_Species_Hierarchy
    'Call Get_Disease_affect_Species
    '
    UserForm_createIN.ListBox_disease.List = Tabel_diseases(1)
    Application.ScreenUpdating = True
End Sub
'This procedure open each table file and extract all relevant columns and save them in a multidimensional array
Public Sub Get_Source_OfInfection()
    '
    Dim wkbook_SOI As Workbook
    '
    Dim sheet_SOI As Worksheet
    '
    'Source of infextion Table Lists
    Dim Lst_SOI_id() As Variant
    Dim Lst_SOI_name() As Variant
    Dim Lst_SOI_is_aquatic() As Variant
    Dim Lst_SOI_is_terrestrial() As Variant
    Dim Lst_SOI_temporal_end_date() As Variant
    Dim Lst_SOI_tr_temporal_end_date() As Variant
    '
    Dim Row_count As Long
    Dim i, arraySize As Integer
    
    '
    Dim Animal_Type As Integer
    ' Get the workbooks of all table files
    Set wkbook_SOI = Workbooks.Open(SourceofInfectionTable_Filename)
    
    ' Get sheets of each workbook
    Set sheet_SOI = wkbook_SOI.Sheets(1)
    
    ' get all relevant columns in Disease table
    With sheet_SOI
    
        'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("b2", .Range("b2").End(xlDown)).Count
        
        'Terristrial disease list
        arraySize = 0
        If UserForm_createIN.OptionButton_Terrestrial Then
            Animal_Type = 0
        Else
            Animal_Type = 1
        End If
    
        For i = 2 To Row_count + 1
            If IIf(Animal_Type = 1, .Cells(i, 4).Value = Animal_Type, Not (.Cells(i, 5).Value = Animal_Type)) And IsEmpty(.Cells(i, 8).Value) And IsEmpty(.Cells(i, 16).Value) Then
                arraySize = arraySize + 1
                '
                ReDim Preserve Lst_SOI_id(arraySize)
                ReDim Preserve Lst_SOI_name(arraySize)
                ReDim Preserve Lst_SOI_is_aquatic(arraySize)
                ReDim Preserve Lst_SOI_is_terrestrial(arraySize)
                ReDim Preserve Lst_SOI_temporal_end_date(arraySize)
                ReDim Preserve Lst_SOI_tr_temporal_end_date(arraySize)
                '
                Lst_SOI_id(arraySize - 1) = .Cells(i, 2).Value
                Lst_SOI_name(arraySize - 1) = .Cells(i, 3).Value
                Lst_SOI_is_aquatic(arraySize - 1) = .Cells(i, 4).Value
                Lst_SOI_is_terrestrial(arraySize - 1) = .Cells(i, 5).Value
                Lst_SOI_temporal_end_date(arraySize - 1) = .Cells(i, 8).Value
                Lst_SOI_tr_temporal_end_date(arraySize - 1) = .Cells(i, 16).Value
           End If
        Next i
        
        ReDim Preserve Lst_SOI_id(arraySize - 1)
        ReDim Preserve Lst_SOI_name(arraySize - 1)
        ReDim Preserve Lst_SOI_is_aquatic(arraySize - 1)
        ReDim Preserve Lst_SOI_is_terrestrial(arraySize - 1)
        ReDim Preserve Lst_SOI_temporal_end_date(arraySize - 1)
        ReDim Preserve Lst_SOI_tr_temporal_end_date(arraySize - 1)
      
        Tabel_ssource_ofInfection() = Array( _
                                            Lst_SOI_id, _
                                            Lst_SOI_name, _
                                            Lst_SOI_is_aquatic, _
                                            Lst_SOI_is_terrestrial, _
                                            Lst_SOI_temporal_end_date, _
                                            Lst_SOI_tr_temporal_end_date)
        Erase Lst_SOI_id
        Erase Lst_SOI_name
        Erase Lst_SOI_is_aquatic
        Erase Lst_SOI_is_terrestrial
        Erase Lst_SOI_temporal_end_date
        Erase Lst_SOI_tr_temporal_end_date
        
    End With
    wkbook_SOI.Close SaveChanges:=False
    
End Sub
'
Public Sub Get_Control_Measures()
    '
    Dim wkbook_CM As Workbook
    Dim wkbook_CM_species As Workbook
    '
    Dim sheet_CM As Worksheet
    Dim sheet_CM_species As Worksheet
    '
    'Contro Measures Table Lists
    Dim Lst_CM_id() As Variant
    Dim Lst_CM_name() As Variant
    Dim Lst_CM_symbol() As Variant
    Dim Lst_CM_is_aquatic() As Variant
    Dim Lst_CM_is_terrestrial() As Variant
    Dim Lst_CM_temporal_end_date() As Variant
    Dim Lst_CM_tr_temporal_end_date() As Variant
    '
    Dim Lst_CM_Domestic() As Variant
    Dim Lst_CM_wild() As Variant
    '
    'SControl measure per animal specie Table Lists
    Dim Lst_CMS_id() As Variant
    Dim Lst_CMS_animal_cat() As Variant
    Dim Lst_CMS_CM_id() As Variant

    '
    Dim Row_count As Long
    Dim i, arraySize As Integer
    Dim animal_type_name, species_domestic, species_wild As String
    '
    Dim Animal_Type As Integer
    ' Get the workbooks of all table files
    Set wkbook_CM = Workbooks.Open(ControlMeasures_Filename)
    Set wkbook_CM_species = Workbooks.Open(ControlMeasures_affect_species_Filename)
    
    ' Get sheets of each workbook
    Set sheet_CM = wkbook_CM.Sheets(1)
    Set sheet_CM_species = wkbook_CM_species.Sheets(2)
    
 
    If UserForm_createIN.OptionButton_Terrestrial Then
        Animal_Type = 0
        animal_type_name = "terrestrial"
        species_domestic = "Domestic"
        species_wild = "Wild"
    Else
        Animal_Type = 1
        animal_type_name = "Aquatic"
        species_domestic = "aquaculture"
        species_wild = "fisheries"
    End If
    '
    With sheet_CM_species
        .AutoFilterMode = False
        Row_count = .Range("a2", .Range("a2").End(xlDown)).Count
        arraySize = 0
        For i = 2 To Row_count + 1
            If InStr(UCase(.Cells(i, 2).Value), UCase(animal_type_name)) > 0 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_CMS_id(arraySize)
                ReDim Preserve Lst_CMS_animal_cat(arraySize)
                ReDim Preserve Lst_CMS_CM_id(arraySize)
                '
                Lst_CMS_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_CMS_animal_cat(arraySize - 1) = .Cells(i, 2).Value
                Lst_CMS_CM_id(arraySize - 1) = .Cells(i, 5).Value

                End If
        Next i
        ReDim Preserve Lst_CMS_id(arraySize - 1)
        ReDim Preserve Lst_CMS_animal_cat(arraySize - 1)
        ReDim Preserve Lst_CMS_CM_id(arraySize - 1)

        '
        Tabel_Control_Measures_affect_species() = Array( _
                                                        Lst_CMS_id, _
                                                        Lst_CMS_animal_cat, _
                                                        Lst_CMS_CM_id)
    End With
    '
    arraySize = 0
    For i = 0 To UBound(Lst_CMS_id)
        If InStr(UCase(Lst_CMS_animal_cat(i)), UCase(species_domestic)) > 0 Then
            arraySize = arraySize + 1
            ReDim Preserve Lst_CM_Domestic(arraySize)
            Lst_CM_Domestic(arraySize - 1) = Lst_CMS_CM_id(i)
        End If
    Next i
    ReDim Preserve Lst_CM_Domestic(arraySize - 1)
    Tabel_Control_Measures_domestic() = Lst_CM_Domestic
    '
    arraySize = 0
    For i = 0 To UBound(Lst_CMS_id)
        If InStr(UCase(Lst_CMS_animal_cat(i)), UCase(species_wild)) > 0 Then
            arraySize = arraySize + 1
            ReDim Preserve Lst_CM_wild(arraySize)
            Lst_CM_wild(arraySize - 1) = Lst_CMS_CM_id(i)
        End If
    Next i
    ReDim Preserve Lst_CM_wild(arraySize - 1)
    Tabel_Control_Measures_wild() = Lst_CM_wild
    
    '
    
   ' get all relevant columns in Disease table
    With sheet_CM
         'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("a2", .Range("a2").End(xlDown)).Count
        arraySize = 0
        'Get all control measures for the Animal Type (Terristriel/Aquatic)
        For i = 2 To Row_count + 1
            If IIf(Animal_Type = 1, .Cells(i, 7).Value = Animal_Type, Not (.Cells(i, 6).Value = Animal_Type)) _
                                And IsEmpty(.Cells(i, 13).Value) And _
                                            IsEmpty(.Cells(i, 24).Value) Then
                arraySize = arraySize + 1
                '
                ReDim Preserve Lst_CM_id(arraySize)
                ReDim Preserve Lst_CM_name(arraySize)
                ReDim Preserve Lst_CM_symbol(arraySize)
                ReDim Preserve Lst_CM_is_terrestrial(arraySize)
                ReDim Preserve Lst_CM_is_aquatic(arraySize)
                ReDim Preserve Lst_CM_temporal_end_date(arraySize)
                ReDim Preserve Lst_CM_tr_temporal_end_date(arraySize)
                '
                Lst_CM_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_CM_name(arraySize - 1) = .Cells(i, 3).Value
                Lst_CM_symbol(arraySize - 1) = .Cells(i, 4).Value
                Lst_CM_is_terrestrial(arraySize - 1) = .Cells(i, 6).Value
                Lst_CM_is_aquatic(arraySize - 1) = .Cells(i, 7).Value
                Lst_CM_temporal_end_date(arraySize - 1) = .Cells(i, 13).Value
                Lst_CM_tr_temporal_end_date(arraySize - 1) = .Cells(i, 24).Value
           End If
        Next i
        
        ReDim Preserve Lst_CM_id(arraySize - 1)
        ReDim Preserve Lst_CM_name(arraySize - 1)
        ReDim Preserve Lst_CM_symbol(arraySize - 1)
        ReDim Preserve Lst_CM_is_terrestrial(arraySize - 1)
        ReDim Preserve Lst_CM_is_aquatic(arraySize - 1)
        ReDim Preserve Lst_CM_temporal_end_date(arraySize - 1)
        ReDim Preserve Lst_CM_tr_temporal_end_date(arraySize - 1)
        
        'Get all control measures for the species (domestic/wild)
      
        Tabel_Control_Measures() = Array( _
                                            Lst_CM_id, _
                                            Lst_CM_name, _
                                            Lst_CM_symbol, _
                                            Lst_CM_is_terrestrial, _
                                            Lst_CM_is_aquatic, _
                                            Lst_CM_temporal_end_date, _
                                            Lst_CM_tr_temporal_end_date)
        Erase Lst_CM_id
        Erase Lst_CM_name
        Erase Lst_CM_symbol
        Erase Lst_CM_is_terrestrial
        Erase Lst_CM_is_aquatic
        Erase Lst_CM_temporal_end_date
        Erase Lst_CM_tr_temporal_end_date
        
    End With
    wkbook_CM.Close SaveChanges:=False
    wkbook_CM_species.Close SaveChanges:=False
    
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
    Dim Lst_disease_temporal_end_date() As Variant
    Dim Lst_disease_translation_name_end_date() As Variant
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
    Dim Animal_Type As Integer
    
    ' Get the workbooks of all table files
    Set wkbook_disease = Workbooks.Open(DiseaseTable_Filename)
    '
    Set wkbook_dcausal = Workbooks.Open(CausalAgentTable_Filename)
    '
    Set wkbook_disease_causal = Workbooks.Open(DiseaseCausalTable_Filename)
    '
       
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
        If UserForm_createIN.OptionButton_Terrestrial Then
            Animal_Type = 0
        Else
            Animal_Type = 1
        End If
    
        For i = 2 To Row_count + 1
            If .Cells(i, 17).Value = Animal_Type And IsEmpty(.Cells(i, 8).Value) And IsEmpty(.Cells(i, 32).Value) Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_disease_id(arraySize)
                ReDim Preserve Lst_disease_name(arraySize)
                ReDim Preserve Lst_disease_group(arraySize)
                ReDim Preserve Lst_disease_OIE_Listed(arraySize)
                ReDim Preserve Lst_disease_Non_OIE_listed(arraySize)
                ReDim Preserve Lst_disease_Emerging(arraySize)
                ReDim Preserve Lst_disease_Self_declaration(arraySize)
                ReDim Preserve Lst_disease_Concern_aquatic(arraySize)
                ReDim Preserve Lst_disease_temporal_end_date(arraySize)
                ReDim Preserve Lst_disease_translation_name_end_date(arraySize)
                '
                Lst_disease_id(arraySize - 1) = .Cells(i, 1).Value
                Lst_disease_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_disease_group(arraySize - 1) = .Cells(i, 9).Value
                Lst_disease_OIE_Listed(arraySize - 1) = .Cells(i, 10).Value
                Lst_disease_Non_OIE_listed(arraySize - 1) = .Cells(i, 11).Value
                Lst_disease_Emerging(arraySize - 1) = .Cells(i, 12).Value
                Lst_disease_Self_declaration(arraySize - 1) = .Cells(i, 15).Value
                Lst_disease_Concern_aquatic(arraySize - 1) = .Cells(i, 17).Value
                Lst_disease_temporal_end_date(arraySize - 1) = .Cells(i, 8).Value
                Lst_disease_translation_name_end_date(arraySize - 1) = .Cells(i, 32).Value

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
        ReDim Preserve Lst_disease_temporal_end_date(arraySize - 1)
        ReDim Preserve Lst_disease_translation_name_end_date(arraySize - 1)
        
        Tabel_diseases() = Array( _
                        Lst_disease_id, _
                        Lst_disease_name, _
                        Lst_disease_group, _
                        Lst_disease_OIE_Listed, _
                        Lst_disease_Non_OIE_listed, _
                        Lst_disease_Emerging, _
                        Lst_disease_Self_declaration, _
                        Lst_disease_Concern_aquatic, _
                        Lst_disease_temporal_end_date, _
                        Lst_disease_translation_name_end_date)
                        
        Erase Lst_disease_id
        Erase Lst_disease_name
        Erase Lst_disease_group
        Erase Lst_disease_OIE_Listed
        Erase Lst_disease_Non_OIE_listed
        Erase Lst_disease_Emerging
        Erase Lst_disease_Self_declaration
        Erase Lst_disease_Concern_aquatic
        Erase Lst_disease_temporal_end_date
        Erase Lst_disease_translation_name_end_date
     
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
        .ComboBox_disease.ListIndex = .ListBox_Diseases.ListIndex
        .Label_disease_group = Tabel_diseases(2)(ListBox_Item_no)
        .Label_OIE_Listed = Tabel_diseases(3)(ListBox_Item_no)
        .Label_Non_OIE_listed = Tabel_diseases(4)(ListBox_Item_no)
        .Label_Emerging_Disease = Tabel_diseases(5)(ListBox_Item_no)
        .Label_Self_Declaration = Tabel_diseases(6)(ListBox_Item_no)
        .Label_is_aquatic = Tabel_diseases(7)(ListBox_Item_no)
        .Label15 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_diseases(0)))
        '
        .ListBox_Causals.List = get_causal_agents(Int(Tabel_diseases(0)(ListBox_Item_no)))
        .TextBox_causal_name.Value = ""
        '
        'UserForm_DiseaseCausal.ListBox_species.List = Tabel_species_group_Terristrial(1)
        Tabel_susceptible_Species = get_Susceptible_Speciese_ids(Int(Tabel_diseases(0)(ListBox_Item_no)))
        .ListBox_species.List = Tabel_susceptible_Species(1)
            
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
    Dim checkAnnimType As Integer
    
    ' Get the workbooks and sheets of all table files and the sheet
    Set wkbook_species_group = Workbooks.Open(SpeciesTable_Filename)
    Set sheet_species_group = wkbook_species_group.Sheets(1)
    '
    With sheet_species_group
        
        'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("h2", .Range("h2").End(xlDown)).Count
        
        'Terristrial disease list
        arraySize = 0
        For i = 2 To Row_count - 1
            checkAnnimType = .Cells(i, 8).Value
            If checkAnnimType = 3 Or checkAnnimType = 4 Or checkAnnimType = 5 Or checkAnnimType = 6 Or checkAnnimType = 7 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_species_group_id(arraySize)
                ReDim Preserve Lst_species_group_name(arraySize)
                ReDim Preserve Lst_species_group_parent_id(arraySize)
                ReDim Preserve Lst_species_group_hierarchy_id(arraySize)
                ReDim Preserve Lst_species_group_enTransl(arraySize)
                '
                Lst_species_group_id(arraySize - 1) = .Cells(i, 3).Value
                Lst_species_group_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_species_group_parent_id(arraySize - 1) = .Cells(i, 7).Value
                Lst_species_group_hierarchy_id(arraySize - 1) = .Cells(i, 8).Value
                Lst_species_group_enTransl(arraySize - 1) = .Cells(i, 14).Value
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
            checkAnnimType = .Cells(i, 8).Value
            If checkAnnimType = 8 Or checkAnnimType = 9 Or checkAnnimType = 10 Or checkAnnimType = 11 Or _
                checkAnnimType = 12 Or checkAnnimType = 13 Or checkAnnimType = 14 Or checkAnnimType = 15 Then
                arraySize = arraySize + 1
                ReDim Preserve Lst_species_group_id(arraySize)
                ReDim Preserve Lst_species_group_name(arraySize)
                ReDim Preserve Lst_species_group_parent_id(arraySize)
                ReDim Preserve Lst_species_group_hierarchy_id(arraySize)
                ReDim Preserve Lst_species_group_enTransl(arraySize)
                '
                Lst_species_group_id(arraySize - 1) = .Cells(i, 3).Value
                Lst_species_group_name(arraySize - 1) = .Cells(i, 4).Value
                Lst_species_group_parent_id(arraySize - 1) = .Cells(i, 7).Value
                Lst_species_group_hierarchy_id(arraySize - 1) = .Cells(i, 8).Value
                Lst_species_group_enTransl(arraySize - 1) = .Cells(i, 14).Value
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
    If UserForm_createIN.OptionButton_Terrestrial Then
        Tabel_species_groups = Tabel_species_group_Terristrial
    Else
        Tabel_species_groups = Tabel_species_group_Aquatic
    End If
    
    
    
    '
    wkbook_species_group.Close SaveChanges:=False

    
End Sub

Private Sub Get_Species_Hierarchy()
    '
    Dim wkbook_species_Hierarchy As Workbook
    Dim sheet_species_Hierarchy As Worksheet
    '
    'Disease Table Lists
    Dim Lst_Hierarchy_id() As Variant
    Dim Lst_Hierarchy_name() As Variant
    Dim Lst_Hierarchy_parent_id() As Variant
    Dim Lst_Hierarchy_AnimalType() As Variant
    Dim Lst_Hierarchy_enTransl() As Variant
    '
    Dim Row_count As Long
    
    ' Get the workbooks and sheets of all table files and the sheet
    Set wkbook_species_Hierarchy = Workbooks.Open(SpeciesHierarchiesTable_Filename)
    Set sheet_species_Hierarchy = wkbook_species_Hierarchy.Sheets(1)
    'ActiveWindow.Visible = False
    '
    With sheet_species_Hierarchy
        
        'Rest Filter
        .AutoFilterMode = False
        Row_count = .Range("a2", .Range("a2").End(xlDown)).Count
        
        
        'Add all columns in disease table to a multidimensional array
        Tabel_species_hierarchy() = Array( _
                                .Range(.Cells(2, 1), .Cells(Row_count + 1, 1)).Value, _
                                .Range(.Cells(2, 2), .Cells(Row_count + 1, 2)).Value, _
                                .Range(.Cells(2, 3), .Cells(Row_count + 1, 3)).Value, _
                                .Range(.Cells(2, 5), .Cells(Row_count + 1, 5)).Value, _
                                .Range(.Cells(2, 8), .Cells(Row_count + 1, 8)).Value)
    End With
    '
    wkbook_species_Hierarchy.Close SaveChanges:=False
    
End Sub


Private Sub ListBox_species_Click()
    '
    Dim ListBox_Item_no, i, s As Integer
    Dim arr_hierarchy() As Variant
    Dim arr_offspring() As Variant
    Dim tree_parent As String
    Dim nParent As Node, nChild As Node
    
    With UserForm_DiseaseCausal
        '
        ListBox_Item_no = .ListBox_species.ListIndex
        arr_hierarchy = get_species_top_parent(Int(Tabel_susceptible_Species(0)(ListBox_Item_no)))
        arr_offspring = get_species_low_offspring(Int(Tabel_susceptible_Species(0)(ListBox_Item_no)))
        '
        .TreeView_hierarchy.LineStyle = 1  'Root lines
        .TreeView_hierarchy.Style = 7
        .TreeView_hierarchy.Nodes.Clear
        If Not IsEmpty(arr_hierarchy(UBound(arr_hierarchy))) Then
            tree_parent = arr_hierarchy(UBound(arr_hierarchy))
        Else
            tree_parent = "NO DATA AVAILABLE"
            'ReDim Preserve arr_hierarchy(UBound(arr_hierarchy))
            arr_hierarchy(0) = tree_parent
        End If
        .TreeView_hierarchy.Nodes.Add Key:=tree_parent, Text:=arr_hierarchy(UBound(arr_hierarchy))
        '
        For i = UBound(arr_hierarchy) To 1 Step -1
            .TreeView_hierarchy.Nodes.Add tree_parent, tvwChild, Key:=arr_hierarchy(i - 1), Text:=arr_hierarchy(i - 1)
            tree_parent = arr_hierarchy(i - 1)
        Next i
        

        For s = 0 To UBound(arr_offspring) - 1
            .TreeView_hierarchy.Nodes.Add tree_parent, tvwChild, Key:=arr_offspring(s), Text:=arr_offspring(s)
            tree_parent = arr_offspring(s)
        Next s


        .Label16 = Str(ListBox_Item_no) & "/" & Str(UBound(Tabel_susceptible_Species(0)))
    End With
End Sub

Function get_top_hierarchy(hierID As Integer) As String
    '
    Dim rowCount, i, top_hierarchy_ID As Integer
    Dim found As Boolean
    '
    rowCount = UBound(Tabel_species_hierarchy(0))
    top_hierarchy_ID = hierID
    For i = 1 To rowCount
         If Tabel_species_hierarchy(0)(i, 1) = top_hierarchy_ID Then
            If IsEmpty(Tabel_species_hierarchy(3)(i, 1)) Then
                Exit For
            Else
                top_hierarchy_ID = Tabel_species_hierarchy(3)(i, 1)
                i = 0
            End If
        End If
    Next i
    get_top_hierarchy = Tabel_species_hierarchy(4)(i, 1) & " (" & Tabel_species_hierarchy(2)(i, 1) & ")"
End Function




Private Sub OptionButton_Aqaua_Click()
    UserForm_DiseaseCausal.CommandButton_get_tables.Enabled = True
End Sub

Private Sub OptionButton_Terre_Click()
    UserForm_DiseaseCausal.CommandButton_get_tables.Enabled = True
End Sub

Private Sub Get_Disease_affect_Species()
    '
    Dim wkbook_disease_affect_Species As Workbook
    Dim sheet_disease_affect_Species As Worksheet
    '
    
    ' Get the workbooks and sheets of all table files and the sheet
    Set wkbook_disease_affect_Species = Workbooks.Open(DiseaseSpeciesTable_Filename)
    Set sheet_disease_affect_Species = wkbook_disease_affect_Species.Sheets(1)
    'ActiveWindow.Visible = False
    '
    With sheet_disease_affect_Species
        
        'Rest Filter
        .AutoFilterMode = False
        
        'Add all columns in disease table to a multidimensional array
        Tabel_disease_affect_Species() = Array( _
                                .Range("b2", .Range("b2").End(xlDown)).Value, _
                                .Range("c2", .Range("c2").End(xlDown)).Value)
        
    End With
    '
    
    wkbook_disease_affect_Species.Close SaveChanges:=False
    
End Sub

Function get_Susceptible_Speciese_ids(diseasID As Integer) As Variant
    '
    Dim rowCount, i, k, arrSize As Integer
    Dim arr_suscp_species_id() As Variant
    Dim arr_suscp_species_names() As Variant
    '
    rowCount = UBound(Tabel_disease_affect_Species(0))
    For i = 1 To rowCount
         If Tabel_disease_affect_Species(1)(i, 1) = diseasID Then
            arrSize = arrSize + 1
            ReDim Preserve arr_suscp_species_id(arrSize)
            ReDim Preserve arr_suscp_species_names(arrSize)
            arr_suscp_species_id(arrSize - 1) = Tabel_disease_affect_Species(0)(i, 1)
            arr_suscp_species_names(arrSize - 1) = get_Speciese_name(Int(Tabel_disease_affect_Species(0)(i, 1)))
         End If
    Next i
    If arrSize > 0 Then
        ReDim Preserve arr_suscp_species_id(arrSize - 1)
        ReDim Preserve arr_suscp_species_names(arrSize - 1)
    Else
        ReDim Preserve arr_suscp_species_id(1)
        ReDim Preserve arr_suscp_species_names(1)
    End If
    get_Susceptible_Speciese_ids = Array(arr_suscp_species_id, arr_suscp_species_names)
End Function

Function get_Speciese_name(speciesID As Integer) As String
    '
    Dim rowCount, arrSize As Integer
    Dim speciesName As Variant
    '
    rowCount = 0
    With UserForm_DiseaseCausal
        '
        Do Until Tabel_species_groups(0)(rowCount) = speciesID Or _
            rowCount = UBound(Tabel_species_groups(0))
             rowCount = rowCount + 1
        Loop
        speciesName = Tabel_species_groups(1)(rowCount)
    End With
    get_Speciese_name = speciesName
End Function

Function get_species_top_parent(speciesID As Integer) As Variant
    '
    Dim rowCount, i, arrSize, top_parent_ID As Integer
    Dim found As Boolean
    Dim arrhierarcy() As Variant
    '
    rowCount = UBound(Tabel_species_groups(0))
    '
    top_parent_ID = speciesID
    '
    arrSize = 0
    '
    For i = 0 To rowCount
         If Tabel_species_groups(0)(i) = top_parent_ID Then
            arrSize = arrSize + 1
            ReDim Preserve arrhierarcy(arrSize)
            arrhierarcy(arrSize - 1) = Tabel_species_groups(1)(i)
            '
            If IsEmpty(Tabel_species_groups(2)(i)) Then
                Exit For
            Else
                top_parent_ID = Tabel_species_groups(2)(i)
                i = 0
            End If
        End If
    Next i
    
    If i > rowCount Then
        i = i - 1
    End If
    
    If arrSize = 0 Then
        ReDim arrhierarcy(0)
    Else
        arrhierarcy(arrSize) = get_top_hierarchy(Int(Tabel_species_groups(3)(i)))
    End If
    
    get_species_top_parent = arrhierarcy
End Function

Function get_species_low_offspring(speciesID As Integer) As Variant
    '
    Dim rowCount, i, arrSize, low_offspring_ID As Integer
    Dim found As Boolean
    Dim ancestry() As Variant
    '
    rowCount = UBound(Tabel_species_groups(0))
    '
    low_offspring_ID = speciesID
    '
    arrSize = 0
    '
    For i = 0 To rowCount
         If Tabel_species_groups(2)(i) = low_offspring_ID Then
            arrSize = arrSize + 1
            ReDim Preserve ancestry(arrSize)
            ancestry(arrSize - 1) = Tabel_species_groups(1)(i)
            low_offspring_ID = Tabel_species_groups(0)(i)
            i = 0
         End If
    Next i
    
    If arrSize = 0 Then
        ReDim ancestry(0)
    End If
   
    get_species_low_offspring = ancestry
    
End Function

Public Function SortThisArray(inputArray As Variant, firstSort As Integer, secondSort As Integer, sortDescending As Boolean) As Variant
    Dim x As String
    Dim y As String
    Dim z As Integer
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    ReDim w(UBound(inputArray, 1), UBound(inputArray, 2))
    If UBound(inputArray, 1) < firstSort Or UBound(inputArray, 1) < secondSort Then
        SortThisArray = inputArray
        Exit Function
    End If
    If firstSort < 0 Then
        firstSort = 0
    End If
    If secondSort < -1 Then
        secondSort = -1
    End If
    If firstSort = secondSort Then
        secondSort = -1
    End If
    For i = 0 To (UBound(inputArray, 2) - 1)
        x = inputArray(firstSort, 0)
        z = 0
        If secondSort <> -1 Then
            y = inputArray(secondSort, 0)
        End If
        For n = 1 To UBound(inputArray, 2)
            If sortDescending = True Then
                If inputArray(firstSort, n) > x Then
                    x = inputArray(firstSort, n)
                    z = n
                    If secondSort <> -1 Then
                        y = inputArray(secondSort, n)
                    End If
                End If
                If inputArray(firstSort, n) = x And secondSort <> -1 Then
                    If inputArray(secondSort, n) > y Then
                        x = inputArray(firstSort, n)
                        z = n
                        y = inputArray(secondSort, n)
                    End If
                End If
            Else
                If inputArray(firstSort, n) < x Then
                    x = inputArray(firstSort, n)
                    z = n
                    If secondSort <> -1 Then
                        y = inputArray(secondSort, n)
                    End If
                End If
                If inputArray(firstSort, n) = x And secondSort <> -1 Then
                    If inputArray(secondSort, n) < y Then
                        x = inputArray(firstSort, n)
                        z = n
                        y = inputArray(secondSort, n)
                    End If
                End If
            End If
        Next n
        For j = 0 To UBound(inputArray, 1)
            w(j, i) = inputArray(j, z)
            inputArray(j, z) = inputArray(j, UBound(inputArray, 2))
        Next j
        ReDim Preserve inputArray(UBound(inputArray, 1), UBound(inputArray, 2) - 1)
    Next i
    For j = 0 To UBound(inputArray, 1)
        w(j, UBound(w, 2)) = inputArray(j, 0)
    Next j
    SortThisArray = w
End Function

Private Sub TreeView_hierarchy_BeforeLabelEdit(Cancel As Integer)

End Sub


Private Sub UserForm_Click()

End Sub


