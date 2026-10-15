let
    // 1. Source Folder se File Read karna (With Safe File & Pending Filter)
    Source = Folder.Files("C:\Users\PC\Desktop\HA_Complaints"),
    #"Filtered Hidden Files1" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true and Text.Contains(Text.Lower([Name]), "pending") and not Text.StartsWith([Name], "~$")),
    #"Sorted Rows" = Table.Sort(#"Filtered Hidden Files1",{{"Date modified", Order.Descending}}),
    #"Kept First Rows" = Table.FirstN(#"Sorted Rows",1),
    #"Invoke Custom Function1" = Table.AddColumn(#"Kept First Rows", "Transform File", each #"Transform File"([Content])),
    #"Renamed Columns1" = Table.RenameColumns(#"Invoke Custom Function1", {"Name", "Source.Name"}),
    #"Removed Other Columns1" = Table.SelectColumns(#"Renamed Columns1", {"Source.Name", "Transform File"}),
    #"Expanded Table Column1" = Table.ExpandTableColumn(#"Removed Other Columns1", "Transform File", Table.ColumnNames(#"Transform File"(#"Sample File"))),
    
    // 2. Data Types aur Column Selection
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Table Column1",{{"Source.Name", type text}, {"COMPLAINT_DATE", type any}, {"CREATION_DATE", type any}, {"ZONE", type text}, {"BRANCH_NAME", type text}, {"PRODUCT", type text}, {"BRAND", type text}, {"CATEGORY", type text}, {"CAPACITY", type text}, {"COMPLAINT_NO", Int64.Type}, {"SERIAL", type text}, {"JOB_TYPE", type text}, {"COMP_WORKING_STATUS", type text}, {"CUSTOMER_MANUAL_NAME", type text}, {"MOBILE_NO", Int64.Type}, {"ADDRESS", type text}, {"EMP_NAME", type text}, {"MODEL_NAME", type text}, {"NO_DAYS", Int64.Type}, {"REMARKS", type text}, {"COMPLAINT_TYPE", type text}, {"LAST_UPDATE_DATE", type datetime}, {"TOTAL_DAYS", Int64.Type}, {"NEW_REMARKS", type text}, {"LOG_REMARKS", type text}, {"PARTS_REQUIRED", type text}, {"ACTIVITY_REMARKS", type text}, {"SSTN", Int64.Type}, {"REMARKS_COMPLAINT_DATE", type date}, {"PROMLEMS", type text}}),
    #"Removed Other Columns" = Table.SelectColumns(#"Changed Type",{"Source.Name", "COMPLAINT_DATE", "BRANCH_NAME", "COMPLAINT_NO", "SERIAL", "JOB_TYPE", "COMP_WORKING_STATUS", "CUSTOMER_MANUAL_NAME", "MOBILE_NO", "ADDRESS", "EMP_NAME", "MODEL_NAME", "NO_DAYS", "REMARKS", "COMPLAINT_TYPE", "ACTIVITY_REMARKS"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Other Columns",{"Source.Name", "COMPLAINT_DATE", "BRANCH_NAME", "JOB_TYPE", "COMPLAINT_NO", "COMP_WORKING_STATUS", "SERIAL", "CUSTOMER_MANUAL_NAME", "MOBILE_NO", "ADDRESS", "EMP_NAME", "MODEL_NAME", "NO_DAYS", "REMARKS", "COMPLAINT_TYPE", "ACTIVITY_REMARKS"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Reordered Columns",{{"COMPLAINT_NO", type text}}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Changed Type1",{"Source.Name", "COMPLAINT_DATE", "BRANCH_NAME", "COMPLAINT_TYPE", "JOB_TYPE", "COMP_WORKING_STATUS", "COMPLAINT_NO", "MODEL_NAME", "SERIAL", "EMP_NAME", "NO_DAYS", "REMARKS", "ACTIVITY_REMARKS", "CUSTOMER_MANUAL_NAME", "MOBILE_NO", "ADDRESS"}),
    
    // 3. Updated System Logic for Current Remarks (Robust Whitespace & Exact Case Matching)
    #"Added Custom" = Table.AddColumn(#"Reordered Columns1", "Current Remarks", each 
        let
            jobType = Text.Trim(Text.From(if [JOB_TYPE] = null then "" else [JOB_TYPE])),
            status = Text.Upper(Text.Trim(Text.From(if [COMP_WORKING_STATUS] = null then "" else [COMP_WORKING_STATUS]))),
            // Handle both UNDER REPAIR and UNDER_REPAIR variants
            isUnderRepair = (status = "UNDER_REPAIR" or status = "UNDER REPAIR"),
            isReRepair = (status = "RE_REPAIR" or status = "RE REPAIR")
        in
            // Priority 1: COMPLETED Status
            if status = "COMPLETED" and jobType = "In Center" then "Ready for Delivery"
            else if status = "COMPLETED" and jobType = "Home Call" then "WAITING FOR OTP"
            
            // Priority 2: In Center (Under Repair / Re Repair)
            else if jobType = "In Center" and (isUnderRepair or isReRepair) then "ARRIVED AT SERVICE CENTER"
            
            // Priority 3: Home Call Under Repair
            else if jobType = "Home Call" and isUnderRepair then "UNDER REPAIR"
            
            // Priority 4: Specific Working Status Rules
            else if status = "PART_WAITING" then "PART_WAITING"
            else if status = "LIFT" then "READY FOR LIFT"
            else if status = "ESTIMATE" then "ESTIMATE"
            else if status = "UNASSIGNED" then "UNASSIGNED"
            else if status = "REJECTED" then "REJECTED"
            
            // Priority 5: Re-repair Logic for remaining cases (e.g. Home Call / General field)
            else if isReRepair then 
                let 
                    txt = Text.Lower(Text.From(if [REMARKS] = null then "" else [REMARKS]) & " " & Text.From(if [ACTIVITY_REMARKS] = null then "" else [ACTIVITY_REMARKS])),
                    isMobility = Text.Contains(txt, "ice") or Text.Contains(txt, "leak") or Text.Contains(txt, "valve") or Text.Contains(txt, "flare") or Text.Contains(txt, "condens") or Text.Contains(txt, "drier") or Text.Contains(txt, "evap") or Text.Contains(txt, "compress") or Text.Contains(txt, "gas") or Text.Contains(txt, "mobility"),
                    isField =  Text.Contains(txt, "testing") or Text.Contains(txt, "fan") or Text.Contains(txt, "smart") or Text.Contains(txt, "pcb collect") or Text.Contains(txt, "flap") or Text.Contains(txt, "tap") or Text.Contains(txt, "pcb") or Text.Contains(txt, "mot") or Text.Contains(txt, "wir") or Text.Contains(txt, "sens") or Text.Contains(txt, "gask") or Text.Contains(txt, "capacit")
                in 
                    if isMobility then "Mobility" 
                    else if isField then "Field" 
                    else "RE_REPAIR (Manual Review)"
            
            // Fallback
            else "In Progress", 
        type text
    ),

    // 4. Self-Referential Step: Safe extraction with robust Type Casting
    ExistingExcelTable = try Excel.CurrentWorkbook(){[Name="HA_Complaints"]}[Content] otherwise null,
    
    HasManualRemarks = ExistingExcelTable <> null and Table.HasColumns(ExistingExcelTable, {"Manual_Remarks"}),
    HasPersonalRemarks = ExistingExcelTable <> null and Table.HasColumns(ExistingExcelTable, {"My_Personal_Remarks"}),
    
    SelectOldRemarks = 
        if ExistingExcelTable = null then
            #table(type table [COMPLAINT_NO = text, Manual_Remarks = any, My_Personal_Remarks = any], {})
        else if HasManualRemarks and HasPersonalRemarks then
            Table.SelectColumns(ExistingExcelTable, {"COMPLAINT_NO", "Manual_Remarks", "My_Personal_Remarks"})
        else if HasManualRemarks then
            Table.AddColumn(Table.SelectColumns(ExistingExcelTable, {"COMPLAINT_NO", "Manual_Remarks"}), "My_Personal_Remarks", each null, type text)
        else if HasPersonalRemarks then
            Table.AddColumn(Table.SelectColumns(ExistingExcelTable, {"COMPLAINT_NO", "My_Personal_Remarks"}), "Manual_Remarks", each null, type text)
        else 
            #table(type table [COMPLAINT_NO = text, Manual_Remarks = any, My_Personal_Remarks = any], {}),

    // Key fix: Dono tables ka COMPLAINT_NO text hona lazmi hai taake mismatch na ho
    EnsureKeyIsText = Table.TransformColumnTypes(SelectOldRemarks, {{"COMPLAINT_NO", type text}}),
    RemoveDuplicates = Table.Distinct(EnsureKeyIsText, {"COMPLAINT_NO"}),

    // 5. Naye Data aur Purane Remarks ko COMPLAINT_NO par Merge karna
    MergeOldRemarks = Table.NestedJoin(#"Added Custom", {"COMPLAINT_NO"}, RemoveDuplicates, {"COMPLAINT_NO"}, "SavedRemarks", JoinKind.LeftOuter),
    
    // 6. Dono columns ko Expand karna
    #"Expanded SavedRemarks" = Table.ExpandTableColumn(MergeOldRemarks, "SavedRemarks", {"Manual_Remarks", "My_Personal_Remarks"}, {"Manual_Remarks", "My_Personal_Remarks"})
in
    #"Expanded SavedRemarks"