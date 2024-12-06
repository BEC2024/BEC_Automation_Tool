Public Class GlobalEntity
    Public Shared dictRawMaterials As Dictionary(Of String, DataSet) = New Dictionary(Of String, DataSet)()

    'Create new part automation form
    '
    'Public Shared Version As String = "1.0.4"

    'Add validation highlight the different thickness in red
    'Public Shared Version As String = "1.0.7"


    'Add the interference report generation
    'Public Shared Version As String = "1.0.32"


    'Update the UI for occurence properties update for selected parts in assembly
    'Public Shared Version As String = "1.0.33"

    'Update the occurence proprty UI, Add the apply colour feature
    '
    'Public Shared Version As String = "1.0.34"

    'Interference use the occurence material
    'Public Shared Version As String = "1.0.35"

    'Inteference new flow apply
    'Check the material property name and set the occurence property interfernce analysis false
    'and generate the interference for all
    'It should generate the interference
    'Public Shared Version As String = "1.0.36"

    ' change Occurance properties in multiple selection
    'Remove readoccurance properties from loadevent
    'Public Shared Version As String = "1.0.37"


    'Update the UI for occurence properties and provide the YES/ NO provision
    'Public Shared Version As String = "1.0.38"

    'Update occurance properties-> Add reference selection 
    ' Public Shared Version As String = "1.0.39"

    'Interfearnce visible true
    'Add wait 
    ' Public Shared Version As String = "1.0.40"

    'Add Interfearnce properties falsebutton and hide getinteferance button
    '  Public Shared Version As String = "1.0.41"

    ' Public Shared Version As String = "1.0.42"

    'Call recursively function for all level 
    ' Public Shared Version As String = "1.0.43"

    'Add Assemblyinterfearnce form 
    'Public Shared Version As String = "1.0.44"

    ' Add inteference top level> on part failed > bug fix
    'Public Shared Version As String = "1.0.45"


    'Merge the interference buttons
    'Public Shared Version As String = "1.0.46"

    'Copy part transfer command open
    'Update all open document command
    'Public Shared Version As String = "1.0.47"

    'add color to copypart
    'Public Shared Version As String = "1.0.48"

    'Update visual studio
    ' Public Shared Version As String = "1.0.49"

    'Create Draftdocument and get partlist For AssemblyBom Tool 
    ' Public Shared Version As String = "1.0.50"

    'change generate report format for child and toplevel,add wait,Checkinteferenceform2 
    'check interference status
    'merge toplevelintereference status report
    ' Public Shared Version As String = "1.0.51"

    'open report location
    'Public Shared Version As String = "1.0.52"

    'MTC Report
    ' Public Shared Version As String = "1.0.53"

    'Filtering in MTC Report,Get properties 
    ' Public Shared Version As String = "1.0.54"

    ' Public Shared Version As String = "1.0.55"

    ' Public Shared Version As String = "1.0.56"

    'Public Shared Version As String = "1.0.57"

    'Add last author column
    ' Public Shared Version As String = "1.0.58"

    'Add color in excel ,change formatting in excel ,add select all button
    'Public Shared Version As String = "1.0.60"

    'Add flatpattern,hardware,doumentno,iscutout validation
    'Public Shared Version As String = "1.0.61"

    'Bug fix : MTC review excel
    'Public Shared Version As String = "1.0.62"

    'Bug fix:
    'Hole cutout Normal cut out
    'Flat pattern
    'Document number logic update
    'Public Shared Version As String = "1.0.63"

    'Add virtualCreationstructure folder
    'Public Shared Version As String = "1.0.64"

    'User Select Assembly and add asm to all level  
    'Public Shared Version As String = "1.0.65"

    'Add authorlist ,check revision level ,check projectname 
    'Public Shared Version As String = "1.0.67"

    'Add occuranceproperties in virtualassemblystructure
    'Public Shared Version As String = "1.0.70"

    'Skip assembly with #name in Virtual Assembly structure
    'Public Shared Version As String = "1.0.71"

    ' Public Shared Version As String = "1.0.72"

    'version73 = vimalbhai
    'Update the single virtual assembly (remove the code and update the excel)

    'Read M2Mfile ,Check title and UOM to M2Mfile ,check dash property 
    ' Public Shared Version As String = "1.0.74"

    'Public Shared Version As String = "1.0.75"
    'Check interferance, add partlistcount,Check SEfeatures, validate comments,upadte physicalproperties
    'Public Shared Version As String = "1.0.76"

    'Public Shared Version As String = "1.0.77"

    ' Check part feature count.
    ' Check assembly feature count.
    'Public Shared Version As String = "1.0.78"

    'CHECK GAGETABLEXCELFILE LINK OR NOT
    'Check issuppressvariable 
    'Public Shared Version As String = "1.0.79"


    'Check sketch is fully defined or not
    'check all-interpartlinks,part-copies,geometrybroken 
    'Public Shared Version As String = "1.0.80"

    'Split the excel with BEC author and other user wise (Create two excels)
    'Public Shared Version As String = "1.0.81"

    'Add the SEO in EXCEL
    'Create the split excel for assembly, part and sheet metal
    'Public Shared Version As String = "1.0.82"

    ' Update baseline category excel of MTC
    ' Create the MTR split data excel 
    'Public Shared Version As String = "1.0.83"

    'Create part sheet-metal
    'Public Shared Version As String = "1.0.84"

    'Mtc report
    'Set color of excel
    'Add thread feature in MTC MTC report
    'Geometry broken feature in MTC MTR report
    'Public Shared Version As String = "1.0.85"

    'Bug fixes for create part.
    'add the provision for file name.
    'Public Shared Version As String = "1.0.86"

    'Update the create part code

    'Public Shared Version As String = "1.0.87"

    'Basline directory path validation added
    'Public Shared Version As String = "1.0.88"

    'Add log in MTC MTR reports
    'Public Shared Version As String = "1.0.89"

    'Merge code for MTC and MTR and create excel for both on single click
    'Public Shared Version As String = "1.0.90"

    'Bug Fix: Mtc MTR report output
    'Public Shared Version As String = "1.0.91"

    'Bug Fix: MTC and MTR report
    'Developer 'Merge the MTC-MTR branch to master branch
    'Public Shared Version As String = "1.0.92"

    'Add Electrical section in Part report
    'Add date in report 
    'Form close on process completed    '
    'Public Shared Version As String = "1.0.93"

    'Add main assembly in report
    'Add silent display alert for skip message
    'Public Shared Version As String = "1.0.94"

    'Check if model is not exists then skip to find features, cutout and recompute
    'Public Shared Version As String = "1.0.95"

    'Change the excel reading of M2m file with  CSV to make reading faster
    'Public Shared Version As String = "1.0.96"

    'Bug fix: read vendor from M2mData
    'Configuration form added
    'Public Shared Version As String = "1.0.97"

    'Bug fix: mtc-mtr report filteration for report not working
    'Public Shared Version As String = "1.0.98"

    'Bug fix filteration
    'Public Shared Version As String = "1.0.99"

    'Update the MTR part> Q-5> hardware part info in report
    'Add the author for all mTR sheets reports > we have added the author as a last questions
    'Part,Baseline MTC report> is adjustable value bug fixed
    'Remove the * logic for file existence as some file include the * in item number
    'Sheetmetal, Part MTC report change the  M2m Source value to Yes, NO and add the details in comment
    'Public Shared Version As String = "1.0.100"

    'Add the modified date for all sheets reports of MTC and MTR report
    'Public Shared Version As String = "1.0.101"

    'For main assembly use the Lastsaved date for Modified > We are not able to find the 'Modified' property
    'Public Shared Version As String = "1.0.102"

    'Skip the invalid sheetmetal part> Some sheet metal part does not have any feature and tool not able to convert that document as sheetmetal part
    'In that case we skip that part and in report show that part as invalid part
    'Public Shared Version As String = "1.0.103"

    'Add the log error message in all try catch.
    'update the cut-out default value to Yes    '
    'Public Shared Version As String = "1.0.104"

    'Update the get revision number code.
    'Add new logic for revision number use last 3 digits with dash
    'Public Shared Version As String = "1.0.105"

    'Bug fix:  Skip the invalid part > part has sheetmetal feature
    'Public Shared Version As String = "1.0.106"

    'Update code for all messages, warning and naming conventions
    'In Previous commit, wrong version name committed
    'Public Shared Version As String = "1.0.107"

    'Add Routing Sequence Report
    'Public Shared Version As String = "1.0.108"

    'Bug fix: Part number is empty in routing sequence report
    '(Document name =partnumber)
    'Bug fix: add format and add the remaining properties
    'Bug Fix: Raw estimation form duplicate path error resolved
    'Add tge hole and bend details in the report sequence report
    '
    'Public Shared Version As String = "1.0.109"

    'Add hole fit, bead gusset,hem details and bend and hole qty for reporting sequence report
    'Public Shared Version As String = "1.0.110"

    'Remove the unit from reporting sequence
    'Public Shared Version As String = "1.0.111"

    'MTC-MTR report sequence excel report> SheetMetalData> bug fix > add louvers,hem etc
    'MTC-MTR report sequence excel report> SheetMetal> Add the part number, Add formula
    'Public Shared Version As String = "1.0.112"

    'Update UI
    'Change form for new part creation
    'Public Shared Version As String = "1.0.113"

    'Bug fix:  AssemblyBOMForm > Integer conversionerror
    'Bug Fix: KPI form
    'Public Shared Version As String = "1.0.114"


    ' In MTC report sequence
    ' Add mass instead of density
    ' Add filepath, material description
    ' Add structure data, structure sheet in reporting sequence report
    ' Add Assembly data, assembly in reporting sequence report
    'Public Shared Version As String = "1.0.115"


    'Update the configuration for raw estimation
    'Update configuration details in all forms
    'Public Shared Version As String = "1.0.116"

    'Set proprity wise holefit 
    '>Set author name in instead of DGS 
    '>Update the excel format
    '> Add new columns like flocation and fbin in excel routing sequence report
    '> Update the pripritywise holefit as per vikas requirement
    '>Use the new bec material excel for add/update part instead of old excel and update code accordingly
    'Public Shared Version As String = "1.0.117"

    '> Add last saved author
    'Public Shared Version As String = "1.0.118"

    'Update routing sequence assembly excel add floc, fbin, qaqc 
    'Update thre report
    'Public Shared Version As String = "1.0.119"

    'added text formate in all routing sequence sheets
    'Public Shared Version As String = "1.0.120"

    'Add qty in all datasheet in Routing sequence report
    'Public Shared Version As String = "1.0.121"


    'Update the last modified date without time for main assembly details in all sheets
    'Update raw material estimation report> update the logic for total length and total order length
    'Public Shared Version As String = "1.0.122"

    'solved part name issue in all data sheets in mtc mtr form
    'Add new category MISC in Routing sequence report
    'Add category field in MISC and Structure report
    'Public Shared Version As String = "1.0.123"


    'remove guideline tab from mainform
    'update assembly validation form on the basis of new bec material excel sheet
    'Raw material BOM excel added in configuration and update code of form(Assembly Bom Form) accordingly
    'Public Shared Version As String = "1.0.124"

    'routing sequence merged 
    'Public Shared Version As String = "1.0.125"

    'routing sequence UI Updated 
    'Public Shared Version As String = "1.0.126"

    'routing sequence>bug fix :solved routing sequence UI issue
    'assembly automation>Bug fix : solved grid issue
    'Public Shared Version As String = "1.0.127"


    'routing sequence> bug fix : added DGV_Combobox_cell value Exact
    'routing sequence > Bug fix : set calculate prodtime in DgvSub
    'mtc mtr report> solved part name date error in routing sequence report 
    'new part creation>added folder path contains  solidedge part instead of debug folder  
    'Update directory name in configuration
    'Public Shared Version As String = "1.0.128"

    'Update output direcory label in all forms
    'Add template path directory in configuration
    'Virtual assembly remove validation for refernce model, allow user to proceed without ref model
    'Update the setup logo and other details
    'Public Shared Version As String = "1.0.129"

    'Bug fix: Virtual structure assembly creation
    'Public Shared Version As String = "1.0.130"

    'virtual structure assembly creation > added checkbox and modify code
    'Public Shared Version As String = "1.0.131"

    'virtual structure assembly creation > added BetterFolderBrowser in Output directory
    'KPI > added manual questions , better folder on browse button and open output dir on complete process 
    'Public Shared Version As String = "1.0.132"

    'config>add RoutingSequenceOutputDirectory
    'Public Shared Version As String = "1.0.133"

    'Routing Sequence:Bug Fix >preview issue solved
    'added help button for all features separately 
    'open folder after complete process in all forms which contains output directory path
    'Routing Sequence >preview form>open solidEdge image on BtnOpenSolidEdge 
    'Public Shared Version As String = "1.0.134"

    'Update setup for help document
    'add 2 new help files routing sequence and configuration
    'Public Shared Version As String = "1.0.135"

    'add log on all forms
    'Public Shared Version As String = "1.0.136"

    'Part/ ShetMetal Update > add Combo-Box Validations
    'Part/ ShetMetal Update > solved Gage Name value while creating Solid-Edge File
    'Assembly Validation > Add Combo-Box Validations
    'created Custom Log Formate
    'adde Custom Log on all Form and add validation for log Directory on Configurtion form
    'Public Shared Version As String = "1.0.137"

    'CustomLogUtil>BugFix:solved ERROR log Message issue
    'Public Shared Version As String = "1.0.138"

    'add solid-edge validation on Below forms
    '1-AssemblyValidation 
    '2-Part_SheetMetalUpdate
    '3-VirtualAssemblyStructure 
    '4-Copy_Transfer_Part
    '5-Occurence 
    '6-RawMaterialEstimation
    '7Interference 
    '8-MTC_MTR_RoutingSeq
    'Public Shared Version As String = "1.0.139"

    'copy&&transfer Part>bugfix:solved solidedge error while opening 
    'added pdf in setup
    'kill Solid-Edge background process in all from
    'Public Shared Version As String = "1.0.140"

    'bugfix:set killSolidedgeprocess on form which create silent solidedge instance 
    'Public Shared Version As String = "1.0.141"

    'BugFix:Solved >New Part Creation form create new part and update lengthwith updation
    'solved issue of formtitle in mainform
    'added : m2m and propseed path from config to mtcmtr form
    'Public Shared Version As String = "1.0.142"

    'added table layout panel in configuration form.
    'set orders in configuration form and set text-align on lables
    'Public Shared Version As String = "1.0.143"

    'New Part Creation:Bug Fix> solved setting variables 
    'Public Shared Version As String = "1.0.144"

    'New Part Creation form:update "Length" lable name to "Height/Length"
    'Public Shared Version As String = "1.0.145"

    'stop kill solidedge process on new part creation form
    'Public Shared Version As String = "1.0.146"

    'Configuration>solved save interferenceExcludeMaterialExcelPath issue 
    'Configuration>Add:MTCExcelPath,MTRExcelPath,RoutingSequenceExcelPath in MTCMTR form
    'Public Shared Version As String = "1.0.147"

    'Part/Sheetmetal Update form : Update Partwise Details
    'Public Shared Version As String = "1.0.148"

    'Part/SheetMetalUpdate : solve error of partwise and materialwise filtering data
    'Part/SheetMetalUpdate : solve Apply button error
    'Public Shared Version As String = "1.0.149"

    'Virtual assembly structure: added customlog
    'Virtual assembly structure: skip error messagebox
    'Interference:skip error messagebox
    'AssemblyAutomation : Added Previous form-> without partwise
    'AssemblyAutomation : update Refresh button code.
    'Public Shared Version As String = "1.0.150"

    'Interference : change date formation in report name.
    'Routing Sequence Tool : update process table if row is already exist it will not add copy in it
    'Public Shared Version As String = "1.0.151"

    'Interference : Apply Property -> IgnoreNonThreadVsThreadConstant
    'Public Shared Version As String = "1.0.152"

    'Interference :Added CustomLog on each button event
    'Public Shared Version As String = "1.0.153"

    'Interference :change report name 
    'Public Shared Version As String = "1.0.154"

    'Interference :change report name after creating the Solidedge Report
    'Public Shared Version As String = "1.0.155"

    'Interference :Removed SolidEdge Report
    'Public Shared Version As String = "1.0.156"

    'Part Creation : Added Template Value for SheetMetal 
    'Part Creation : Added Gage Name Property in Custom Property for SheetMetal
    'Interference :Added try and catch for solidedge CheckInterference2 method and activate report for that.
    'Public Shared Version As String = "1.0.157"

    'Interference :in checkinterference2 set AddOccurance=false
    'this changes are working in Clients's system , it was issue in client's system that could not add part document in assembly after excecution
    'Public Shared Version As String = "1.0.158"

    'Interference :To False Interferance Occurance Properties>>Solved Issue of Could not Generating Report and Adding Occurance part Names 
    'Public Shared Version As String = "1.0.159"

    'Interference :To False Interferance Occurance Properties>>Added FolderBrowserDialog for OutputDirectory Path
    'Public Shared Version As String = "1.0.160"

    'Routing Sequence : Remove Duplicate Orders While execute Apply Value Button.
    'Public Shared Version As String = "1.0.161"

    'Part Creation : update code for Diameter that contains (ROUND OR TUBING)
    'Public Shared Version As String = "1.0.162"

    'Copy and Transfer part : Solved Error
    'Public Shared Version As String = "1.0.163"

    'New Part Creation : Replace Length to height
    'New Part Creation : Add new Feature Linear length, add column in excel,Filter on Type and Bec Code/Material Used for Linear Length
    'Public Shared Version As String = "1.0.164"

    'New Part Creation: Set Filter of data for Structure sheet into form properties
    'New Part Creation: set color for GageNme property if it is "Not Bendable"
    'Public Shared Version As String = "1.0.165"

    'New Part Creation : Solve the Filter issue in the Structure category
    'Public Shared Version As String = "1.0.166"

    'New Part Creation : Add Highlight to textboxes which contains blank,Missing or Not values
    'PartSheetMetalUpdate:hide PartWise and Materialwise radio button, bydefault get Materialwise data
    'Public Shared Version As String = "1.0.167"

    'temp7APR2023
    'PartSheetMetalUpdate:Add Filter to the Gage Name Propety.
    'Public Shared Version As String = "1.0.168"

    'temp11APR2023
    'Public Shared Version As String = "1.0.169"
    'VirtualAssemblyStructure:Update the flow according to the Client's Excel
    'VirtualAssemblyStructure:in Top And Main Level Assembly Added Title Property of SummaryInformation
    'VirtualAssemblyStructure:Remove Add User Assmebly And Reference Model Feature(Visible false)

    'temp13APR2023
    'Public Shared Version As String = "1.0.170"
    'AssemblyValidations:Change the Design and add both side fields into tablelayoutpannel
    'AssemblyValidations:Add new Combobox  BendType in Materialwise 
    'AssemblyValidations:add new validations in all current Textboxes that contains blank,Missing or Not values.

    'Temp18APR2023
    'PartSheetMetalUpdate:added validation for bend-radius if it contrast with current Bend-radius.
    'Use FastExcel dll for reading BECMaterial.xlsx file
    'Solve error of filter in New Part craetion,PartSheetMetal And Assembly Validation form
    'Public Shared Version As String = "1.0.171"

    'Temp24APR2023
    'PartSheetMetalUpdate:Solved Auto Suggest According to Current Material Used Property.
    'Public Shared Version As String = "1.0.172"

    'TEMP25APR2023
    'PartSheetMetalUpdate:Added Auto Suggest Same for Gage Name if it is exit in Current Field.
    'Public Shared Version As String = "1.0.173"

    'TEMP27APR2023
    'PartSheetMetalUpdate:Solve Issue of AutoSuggest for BendType ,SetGageName And Refresh Button
    'Public Shared Version As String = "1.0.174"

    'TEMP27APR2023
    'Public Shared Version As String = "1.0.175"
    'PartSheetMetalUpdate:Solve Issue of Refresh Button

    'TEMP02MAY2023
    'Comment MTR Code and Rename MTC/MTR to MTC
    'RawMaterialEstimation:Added BEC Material Excel Instead of RawMaterialBOM Excel For Fetching Data
    'Made Changes in RawMaterialEstimation According client's excel Formate
    'Public Shared Version As String = "1.0.176"

    'RawMaterialEstimation:
    'Added logic for MaterialUsed if it has contain width and height
    'added Notfound values for Category and StandardThickness
    'Public Shared Version As String = "1.0.177"

    'added Validation in RawMaterialEstimation in Generate Report
    'Public Shared Version As String = "1.0.178"

    'Added Validations for NF and N/A
    'Public Shared Version As String = "1.0.179"

    'RawMaterialEstimation->
    '1 Added Order arrangement , Not found  Values validation  in Sheet-plate-structure sheet 
    '2 Validation in Sheet-Palte-Structure,Std-Parts-And-Hardware, Misc Table
    'Public Shared Version As String = "1.0.180"

    'RawMaterialEstimation->Add Order Sequence in Std parts and Misc datatable too
    'Public Shared Version As String = "1.0.181"

    'RawMaterialEstimation->fix the Total height,length and other values based on height and length
    'Public Shared Version As String = "1.0.182"

    'MTC Report
    '(1)Remove the MTR test from configurations
    '(2)Remove strikes-out questions from the MTC  form
    '(3)Remove strikes-out questions from the MTC  form
    'Public Shared Version As String = "1.0.183"

    'KPI Report : Remove the MTC Name From Lable
    'RawMaterial Estimation : Replace BEC Number to BEC Material Code
    'Public Shared Version As String = "1.0.184"

    'Raw Material Excel : Change the Sheet Name of Generated Excel Report
    'Public Shared Version As String = "1.0.185"

    'change position of textbox in config file
    'Public Shared Version As String = "1.0.186"

    'change the flow of virtual structure
    'added group in virtualstructure code
    'Public Shared Version As String = "1.0.187"

    'Added group and Title in Virtual Assembly
    'Public Shared Version As String = "1.0.188"

    'Added Login in config form
    'Added Changing options for ConfigProperties.xml 
    'Public Shared Version As String = "1.0.189"

    'Added Validation in Part ShetMetal Update
    'Public Shared Version As String = "1.0.190"

    'Added BEC_Automation_Installer for removing BEC Folder From AppData after Uninstall of setup
    'Changes in Config form: visible false Raw Material BOM Excel Path
    'PartSheetMetalUpdate :set auto Suggest
    'Assembly Validation:Update Filter Issue
    'All Form:Visible false button and enable false input and output textbox 
    'Public Shared Version As String = "1.0.191"

    'RAW Material Estimation: Replace Config RAW Material BOM path to BEC Material Excel path
    'Public Shared Version As String = "1.0.192"

    'PartSheetMetal Update:Added Table Layout and All Properties set as Organised Manners
    'Public Shared Version As String = "1.0.193"

    'MTC Form: Solve filter issue in BEC and DGS Assembly Sheet
    'MTC Form: Open Active Document after report generate.
    'Public Shared Version As String = "1.0.194"


    'MainForm:
    '(1)on Form Activated event added Auto Save event of Solidedge 
    '(2)Set Author Updation of Summary Information of activated Solid Edge Document
    'Config Form
    '(1)Added AutoSaveAuthor check box 
    'MTC Form
    '(1) Filter Title of SolidEdge Document till left 25 Charaters with M2M file.
    'Public Shared Version As String = "1.0.195"

    'Public Shared Version As String = "1.0.196"

    'Public Shared Version As String = "1.0.197"

    'BOMCount Issue Solved
    'Public Shared Version As String = "1.0.198"

    'MTC : partlist count manage fucntionwise - mainassemblydata and subassmeblydata
    'MTC : DGS And BEC Issue Solved 
    'Public Shared Version As String = "1.0.199"

    'PartSheetMetalUpdate - Enable false Size ,Part type , Garde
    'New Part Creation - GageName column set as per priority of Bec Excel.
    'Public Shared Version As String = "1.0.200"

    'New Part Creation - Filter priority to lowest
    'Assembly Validation - add try catch on  GetMaterialLibraryLists to stop crashing tool. 'TEMP_6SEPT203
    'Public Shared Version As String = "1.0.201"

    'TEMP12SEPT2023
    'PartSheetMetalUpdate: Added bendradius from variable table in Current bendradius textbox
    'AssemblyUpdate: Added bendradius from variable table in Current BendRadiusTextbox
    'PartSheetMetalUpdate: Color validate for Bend Radius on Current BendRadiusTextbox
    'AssemblyUpdate: Color validate for Bend Radius on Current BendRadiusTextbox
    'Public Shared Version As String = "1.0.202"

    'PartSheetMetalUpdate: in Variable Table Set RadiusGloble and BendRadius to 1st and 2nd priority
    'Assembly Validation: in Variable Table Set RadiusGloble and BendRadius to 1st and 2nd priority
    'Public Shared Version As String = "1.0.203"

    'PartSheetMetalUpdate:  Filter priority to lowest in bendType
    'Assembly Validation:  Filter priority to lowest in bendType
    'Public Shared Version As String = "1.0.204"

    'Assembly Validation:Bend Radius filter issue solved
    'MTC:In Excel report Yes/No Rule Correction in SheetMetal question 8 and 18 
    'config form : visible  Set Configuration Path on login from too
    'config form : solve issue on  "Set Configuration Path" button
    'Config From:Refresh from after click on button "Set Configuration Path"
    'Assembly Validation: set layout of Input-Boxes
    'Public Shared Version As String = "1.0.205"

    'TEMP27-SEP-2023
    'New Part Creation:Add SummaryInfoProperty function for filling Author Name
    'Raw MAterial Estimate : Solved excel report columns( Length , Width ) NotFound issue

    'TEMP29SEP2023
    'Raw Material Estimate : Solve Double convertion issue in Length , Width , Total Length ,Total Width,Area,Order Area 
    'Raw Material Estimate : Merege issue solved in Excel
    'KPI : Solve Error of Generating Report
    'MTC Report: Add changes in design,change pdf name to "MTC Help.pdf" date:25/10/2023
    'Public Shared Version As String = "1.0.206"


    'Public Shared Version As String = "1.0.207"

    'temp 2Feb2024
    'Added a dropdown list For "Gage Name" in "Part/Sheet-Metal Update" form.
    'Also update dropdown list For "Gage Name" As per "Material Used" list, whenever selection in "Material Used" dropdown list changed.
    'Public Shared Version As String = "1.0.208"


    'temp 6Feb2024
    'Added a code in "Part/Sheet-Metal Update" form.
    'Update value for "Bend Radius" as per "Gage Name" list in textbox "txtBendRadius_Mw", whenever selection in "Gage Name" dropdown list changed.
    'Public Shared Version As String = "1.0.209"

    'temp 20Feb2024
    'Added code to merge both excel files "BEC_MTC Report" and "DGS_MTC Report". 
    'Updated code to manage, formatting for the findings column.
    'Updated Excel File "MTC_BEC.xlsx" as per client requirements.
    'Public Shared Version As String = "1.0.210"

    '2nd Sep 2024
    'MTC Report: Align all text to the center and middle
    'MTC Report: Set the default tab opens to the "Assembly" tab, which should always be the first tab
    'MTC Report: Combine the BEC and DGS reports into one
    'Public Shared Version As String = "1.0.211"

    '2nd Sep 2024
    'Routing Sequence Report: Save the Routing sequence report on a individual path
    'Routing Sequence Tool: Remove this tool entirely
    'Routing Sequence Tool: Deleted button 'BtnQCRoutingSequence', which directs to form 'RST_Design1'
    'Public Shared Version As String = "1.0.212"

    '5th Sep 2024
    'Template File: Make the report Excel file a template (.xltx)
    'Update: Baseline Sheet, Remove marked lines and update code accordingly
    'Update: UI Alignment in Configuration Form
    'Author Information: Add the MTC author’s name and a timestamp in the first row of every tab in MTC report
    'Issue Addressed Column: Include an additional column beside every part number in the report, with the header "Issue Addressed? YES / NO / Remark"
    'KPI Tool Issue: Fix the reported issue in the KPI tool 
    '9th Sep 2024
    'MTC Report: Updates in excel templates
    '13th Sep 2024
    'Update: New MTC_BEC Excel Template
    'Revision Number Validation: Address the validation issue
    'Text Alignment: Align all text to the center and middle (repeat for emphasis)
    'Public Shared Version As String = "1.0.213"


    '17th Sep 2024
    'MTC Report: Added code to set Revision Number as Correct if Revision Number's value is '0'
    'MTC Report: Added code to generate MTC Report for individual Part (.par) files
    'Public Shared Version As String = "1.0.214"

    '17th Sep 2024
    'MTC Report: Added code to generate MTC Report for individual SheetMetal (.psm) files
    'Public Shared Version As String = "1.0.215"

    '19th Sep 2024
    'Bug Fix: Configuration Error, Pop-up message remains in loop (Employee excel file missing or not exist in config)
    'Public Shared Version As String = "1.0.216"

    '26th September 2024
    'Updates in MTC reports for Assembly (UOM Property)
    'Updates in MTC reports for Baseline (Is Hole Tool Used)
    'Updates in MTC reports for Baseline (Hardware Part)
    'Updates in MTC reports for Baseline (UOM Property)
    'Updates in MTC reports for Baseline (match with M2M)
    'Updates in MTC reports for Baseline (Is Valid Directory Path)
    'Public Shared Version As String = "1.0.217"

    '30th September 2024
    'Bug Fix: Raw Material Estimation
    'Added 'BEC RAW' Draft list
    'Public Shared Version As String = "1.0.218"

    'Created custom setup of BEC project and publish it on git 
    'Changed excel name format for Raw Material Estimation file
    'Public Shared Version As String = "1.0.219"

    '10th october 2024
    'Change UI for 'Change/Save Config'
    'Replace Help Documents with new pdfs
    'Public Shared Version As String = "1.0.220"

    '14th october 2024
    'Configuration Form UI: Added 'Help Document Directory'
    'Added 'Help Document Directory' element in config file
    'Added code to place all 'Help Documents' from installation path, if user changes path for 'Help Document Directory'
    'Public Shared Version As String = "1.0.221"

    '16th october 2024
    'Sheetmetal part update tool - turn off/disable the drop-down for the 'Gage Name' field
    'Update the path of help documents on 'Help Button'
    'Public Shared Version As String = "1.0.222"

    '28th october 2024
    'MTC Report (Baseline): Update for SE Simplified Features
    'MTC Report: Add User name who creates the report
    'MTC Report: Remove 'Remark' from the dropdown and keep only 'YES' and 'NO'
    'MTC Report: Add an extra column for each part, titled 'Remark'
    'Remove 'Routing Sequence Directory' from 'Configuration' and 'MTC' forms
    'Remove 'MTC' word from 'MTC Report'
    'Match label titles with configuration's titles to ensure consistency
    'Remove extra '&' from the file name in the tool code
    'Remove 'MTR' from the file name in the tool code
    'Update the KPI tool to match the latest MTC output report file 
    'Update 'BEC-Material' excel in excel templates
    'Update Color format for MTC excel report
    'Public Shared Version As String = "1.0.223"

    '4th Nov 2024
    'KPI Tool: Update code of KPI Tool as per new format of MTC Excel Reports
    'Public Shared Version As String = "1.0.224"

    '5th Nov 2024
    'Bug-Fix:  Fetch the correct values for 'Gage Name' and 'Bend Type' in the Part/Sheet-Metal Update Form
    'Public Shared Version As String = "1.0.225"

    '6th Nov 2024
    'Bug-Fix: KPI report not in string format
    'Public Shared Version As String = "1.0.226"

    '8th Nov 2024
    'Bug-Fix: MTC report cell formatting
    'Public Shared Version As String = "1.0.227"

    '8th Nov 2024
    'Remove subfolder creation when tool run on individual Part and Sheetmetal
    'Add MTC author name (window's user) in Part and Sheetmetal report instead of Model author name
    'Public Shared Version As String = "1.0.228"

    '15th Nov 2024
    'Removed MoCustom Control Project
    'Public Shared Version As String = "1.0.229"

    '19th Nov 2024
    'MTC Report: tool is not identifying a baselined status part for individual model reports
    'MTC Report: tool is not identifying a electrical status part for individual model reports
    'Public Shared Version As String = "1.0.230"

    '21st Nov 2024
    'MTC Report: Identify individual part status, if it's 'baseline' or 'electrical' then general MTC Report accordingly
    '(if it's baseline then generate report of baseline sheet, if it's elecrical then generate report of electrical sheet, otherwise generate report of part sheet)
    'Public Shared Version As String = "1.0.231"

    '25th Nov 2024
    'Bug Fix: MTC Report - fetch exact value of 'Title' property
    'Public Shared Version As String = "1.0.232"

    '25th Nov 2024,28th Nov 2024 and 29th Nov 2024
    'Code update: Part number does not match with M2M
    'Part number M2M : compare model document number with fPart from RPIMAS file
    'Code update: Model title does not match with M2M item master description field
    'Model title M2M : compare model title with fDescription from RPIMAS file
    'Public Shared Version As String = "1.0.233"

    '2nd Dec 2024
    'Everyehwere added the partnumber question value to document correct value
    '{mtcAssemblyObj.documentNumberCorrect = mtcAssemblyObj.partNumber} -- added
    '{If((mtcReviewObj.fileNameWithoutExt.Contains(mtcReviewObj.documentno)), "Yes", "No") -- removed
    'Public Shared Version As String = "1.0.234"

    '3rd Dec 2024
    'Bug Fix: MTC Report Baseline Sheet (6. Are ALL features fully defined in the model Pathfinder? * (Sketch fully constraint?))
    Public Shared Version As String = "1.0.235"
End Class
