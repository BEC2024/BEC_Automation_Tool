Public Class RountingSequenceClass
    Public i As Integer = Nothing
    Public excelFilepath As String = Nothing
    Public dir As String = Nothing

    Public PartName As String = Nothing
    Public user As String = Nothing
    Public MaterialDescription As String = Nothing
    Public FilePath As String = Nothing
    Public image As Image = Nothing
    Public order As ArrayList = New ArrayList()
    Public prodtime As ArrayList = New ArrayList()
    Public movetime As ArrayList = New ArrayList()
    '----------------------------------------------------
    Public ProjectName As String = Nothing
    Public CategoryName As String = Nothing
    ' Public partname
    Public ten = "9020"
    Public twenty = "9130"
    Public thirty = "1020"
    Public fourty = "1040"
    Public fifty As String
    Public fifty1 = "3010"
    Public fifty2 = "2090"
    Public fifty3 = "2040"
    Public fifty4 = "2100"
    Public sixty As String
    Public sixty1 = "3030"
    Public sixty2 = "3035"
    Public seventy = "7020"
    Public eighty = "7050"
    Public ninety = "2030"
    Public hundred = "2035"

    Public process_ten = "Material Handling"
    Public process_twenty = "Nesting"
    Public process_thirty = "Cutting Center"
    Public process_fourty = "Grind/Buff"
    Public process_fifty1 = "Layout & Prep (Pickering St.)"
    Public process_fifty2 = "Manual Mills"
    Public process_fifty3 = "Doosan CNC Mill"
    Public process_fifty4 = "Radial Arm Drill"
    Public process_sixty1 = "Brake, Press 400 Ton"
    Public process_sixty2 = "Brake, Press 240 Ton (Pickering St.) "
    Public process_seventy = "Metal Preparation"
    Public process_eighty = "Paint/Undercoat"
    Public process_ninety = "Lathe Summit"
    Public process_hundred = "Lathe Jet 12X36"
    Public dtProcess As DataTable = New DataTable

    Public Maindt As DataTable = New DataTable
    Public dt2 As DataTable = New DataTable
    Public dt3 As DataTable = New DataTable
    Public dgvmainRowIndex As DataGridViewRow
    Public dgvmainRowIndexProdTIME As DataGridViewRow
    Public dgvmainRowIndexMoveTIME As DataGridViewRow
    Public dgvCalculatorRowIndex As DataGridViewRow
End Class
