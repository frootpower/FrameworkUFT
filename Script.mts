
Dim MRowCount, i, ModuleExe, ModuleId, TCRowCount, j, ModuleId2, TestCaseExe, TestCaseId
Dim TSRowCount, k, TestCaseId2, Keyword
Dim almExecExcelPath
Dim almExecPath, almCycleName, almTestCaseID, bpmScenario
Dim almExecPathArr

'getAUTRootFolder = Pathfinder.locate("..\..\")
'Envvarpath = getAUTRootFolder & "Data\Environment.xml"
'organizerPath = getAUTRootFolder & "Organizer\Organizer_New.xls"
'url = environment.Value("URL")

'Environment.LoadFromFile(Envvarpath)

getConfigurationValue()
'msgbox environment.Value("URL")

'Add new sheet to Run-time Data Table to import instructions from the Organizer
DataTable.AddSheet "Module"
DataTable.AddSheet "TestCase"

'Import data from an external file
'DataTable.ImportSheet "C:\Users\keshubm\WORK\UFT\Automation Framework\Final Automation Framework\Project_Framework_V1\Organizer\Organizer_New.xls", 1, 3
'DataTable.ImportSheet "C:\Users\keshubm\WORK\UFT\Automation Framework\Final Automation Framework\Project_Framework_V1\Organizer\Organizer_New.xls", 2, 4
'DataTable.ImportSheet "D:\Project_Framework\Organizer\Organizer_New.xls", 3, 5

DataTable.ImportSheet organizerPath, 1, 3
DataTable.ImportSheet organizerPath, 2, 4

'Read executable Module ids from Module sheet
MRowCount= DataTable.GetSheet(3).GetRowCount

For i = 1 To MRowCount Step 1
    DataTable.SetCurrentRow(i)
    
    ModuleExe = DataTable(3, 3)
    
    If UCase(ModuleExe) = "Y" Then
        ModuleId = DataTable(1, 3)
        'Msgbox ModuleId
        
        
'Read executable Test Case id's under executable modules from Test Case sheet
'************************************ previous ***********************
'TCRowCount = DataTable.GetSheet(4).GetRowCount
'For j = 1 To TCRowCount Step 1
'    DataTable.SetCurrentRow(j)
'    ModuleId2 = DataTable(4, 4)
'    TestCaseExe = DataTable(3, 4)
'    
'    If UCase(TestCaseExe) = "Y" And  ModuleId = ModuleId2 Then
'        TestCaseId = DataTable(1, 4) 
'       ' Msgbox TestCaseId
'        
'        fnReadExcel(TestCaseId)
''SAPGuiSession("Session").SAPGuiWindow("SAP").Resize 218,26
'        
'
''RunAction "Copy of Action1", oneIteration
'
''		RunAction "Action1 [Create_Product_Requisition]", oneIteration
'   End If
'  Next
'End If
'Next
'
'**********************************Later**********************************
TCRowCount = DataTable.GetSheet(4).GetRowCount
For j = 1 To TCRowCount Step 1
    DataTable.SetCurrentRow(j)
    almExecExcelPath = DataTable(3, 4)
    TestCaseExe = DataTable(4, 4)
    ModuleId2 = DataTable(5, 4)
    bpmScenario = DataTable(6, 4)
    If UCase(TestCaseExe) = "Y" And  ModuleId = ModuleId2 Then
    
        almExecPathArr = dataProvider(almExecExcelPath)
        almExecPath = almExecPathArr(0)
        almCycleName = almExecPathArr(1)
        almTestCaseID = almExecPathArr(2)
        
        Call exec_HP_ALM(almExecPath, almCycleName, almTestCaseID)
        
        configureExecutionSheetName(bpmScenario)
        
        TestCaseId = DataTable(1, 4) 
       ' Msgbox TestCaseId
        
        fnReadExcel(TestCaseId)
'SAPGuiSession("Session").SAPGuiWindow("SAP").Resize 218,26
        

'RunAction "Copy of Action1", oneIteration

'        RunAction "Action1 [Create_Product_Requisition]", oneIteration
   End If
  Next
End If
Next



'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Maintain Outb\. Deliv\. Order - Warehouse No\. KZ01 \(Time Zone EST\)").SAPGuiComboBox("type:=GuiComboBox","attachedtext:=Find").Select("ERP Document")
'
'Dim mySendKeys,ObjSAPGuiTree,ObjKeyValues,arrNumbers(),ERPDeliverynumber,intCount, OutboundDeliverynumber, WarehouseOrderNumber @@ hightlight id_;_1442886_;_script infofile_;_ZIP::ssf2.xml_;_
'intCount=0
'SAPGuiUtil.AutoLogon "ECQ - SYK S4H QA", 900, "kmathur", "#1iskcon", "EN"
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/nVA01"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiButton("type:=GuiButton","tooltip:=Enter").Click
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Sales Document").SAPGuiEdit("attachedtext:=Order Type","Index:=0").Set "ZKB"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Sales Document").SAPGuiEdit("attachedtext:=Sales Organization","Index:=0").Set "1710"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Sales Document").SAPGuiEdit("attachedtext:=Distribution Channel","Index:=0").Set "10"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Sales Document").SAPGuiEdit("attachedtext:=Division","Index:=0").Set "20"
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Sales Document").SAPGuiButton("name:=btn\[0\]").Click
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiEdit("attachedtext:=Sold-To Party", "Index:=0").Set "210936"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiEdit("attachedtext:=Ship-To Party", "Index:=0").Set "210936"
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiButton("name:=btn\[0\]").Click
'
'''SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Partner selection").SAPGuiLabel("content:=FedEx Express").SetFocus	'This has been commented.
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Partner selection").SAPGuiButton("name:=btn\[0\]").Click
'wait(5)
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiTable("name:=SAPMV45ATCTRL_U_ERF_AUFTRAG").SetCellData 1, "Material", "53-34804"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiTable("name:=SAPMV45ATCTRL_U_ERF_AUFTRAG").SetCellData 1, "Order Quantity", "1"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiTable("name:=SAPMV45ATCTRL_U_ERF_AUFTRAG").SetCellData 1, "Plnt", "0003"
''
''
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiButton("name:=btn\[11\]").Click
'''SAPGuiSession("Session").SAPGuiWindow("Create SYK Consign. Fill").Maximize
''
'''SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr\.pr\.: Variants").SAPGuiToolbar("name:=GridToolbar").PressButton "IMMSTART"
''
'''val = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiGrid("micclass:=SAPGuiGrid","title:=Variants for Program RVV50R10C").Object(
'''msgbox val
'''
'''SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr\.pr\.: Variants").SAPGuiButton("micclass:=SAPGuiButton","text:=Start immediately").Click
'''
'''SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiGrid("micclass:=SAPGuiGrid","title:=Variants for Program RVV50R10C").Object.SelectContextMenuItemByText(val)
'''SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiGrid("micclass:=SAPGuiGrid","title:=Variants for Program RVV50R10C").Object.ClickCurrentCell()
'''
'''
'''SAPGuiSession("Session").SAPGuiWindow("Create deliveries in bckgr.pr.").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf15.xml_;_
'
'salesOrderNum = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiStatusBar("micclass:=SAPGuiStatusBar").GetROProperty("item2")
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/nVL10BATCH"
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign. Fill Up: Overview").SAPGuiButton("name:=btn\[0\]").Click
'
'rowval = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr.pr.: Variants").SAPGuiGrid("title:=Variants for Program RVV50R10C").FindRowByCellContent("Variant Name","Z_0003_FEDEXP")
'
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr.pr.: Variants").SAPGuiGrid("title:=Variants for Program RVV50R10C").SelectRow(rowval)
'''SAPGuiSession("Session").SAPGuiWindow("Create deliveries in bckgr.pr.").SAPGuiToolbar("GridToolbar").PressButton "IMMSTART"		This is a comment @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf15.xml_;_
'
'Set objToolbar1 = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr.pr.: Variants").SAPGuiToolbar("title:=Variants for Program RVV50R10C")
'maxIndex = objToolbar1.Object.ToolbarButtonCount 
'msgbox maxIndex
'strTooltip = "Start immediately"
'For i = 0 To maxIndex
'   If strTooltip = objToolbar1.Object.GetToolbarButtonTooltip(i) then
'       strButtonID = objToolbar1.Object.GetToolbarButtonId(i) 
'       msgbox strButtonID
'       objToolbar1.Object.Presstoolbarbutton strButtonID
'       Exit For
'   End if
'Next
'
'SAPGuiSession("Session").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr.pr.: Variants").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/nVA03"
'SAPGuiSession("Session").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create deliveries in bckgr.pr.: Variants").SAPGuiButton("name:=btn\[0\]").Click
'
'msgbox salesOrderNum
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Display Sales Document").SAPGuiEdit("attachedtext:=Order").Set salesOrderNum
'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Display Sales Document").SAPGuiButton("name:=btn\[17\]").Click
'
'''Extracting the delivery number
'Set ObjSAPGuiTree = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Document Flow").SAPGuiTree("micclass:=SAPGuiTree").Object
'Set ObjKeyValues = ObjSAPGuiTree.GetAllNodeKeys
'
'''get the total count
'''This count indicates the number of items/nodes in the Tree
'intNodeCount = ObjKeyValues.Count
'
'For index = 0 to intNodeCount-1
'    ''Get the node text
'    strNodeText=ObjSAPGuiTree.GetNodeTextByKey(ObjKeyValues(index))
'    
'    If Instr(strNodeText,"Delivery")>0 Then
'    	msgbox strNodeText
'    	
'    	j=0			''starting index for Numbers array.
'		''	msgbox "Getting Numbers"
'		splitStr = Split(strNodeText, " ")
'		intCount = UBound(splitStr)
'		''intCount = Cint(intCount)
'		msgbox intCount
'		For i = 0 To intCount 
'			If IsNumeric(splitStr(i)) Then
'				ReDim preserve arrNumbers(j)
'				arrNumbers(j) = splitStr(i)
'				''			msgbox arrNumbers(j)
'				j = j + 1
'			End If
'		Next
'    	ERPDeliverynumber = arrNumbers(0)
'    	msgbox ERPDeliverynumber
'    	Exit For
'    End If
'    
'Next
'
'SAPGuiUtil.AutoLogon "EMQ - SYK EWM QA", 900, "kmathur", "#1iskcon", "EN"		'Logging into SAP EWM
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/n/scwm/prdo"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiButton("name:=btn\[0\]").Click
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Maintain Outb\. Deliv\. Order - Warehouse No\. KZ01 \(Time Zone EST\)").SAPGuiComboBox("type:=GuiComboBox","attachedtext:=Find").Select("ERP Document")
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Maintain Outb\. Deliv\. Order - Warehouse No\. KZ01 \(Time Zone EST\)").SAPGuiEdit("type:=GuiTextField","attachedtext:=Find").Set ERPDeliverynumber
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Maintain Outb\. Deliv\. Order - Warehouse No\. KZ01 \(Time Zone EST\)").SAPGuiButton("type:=GuiButton","tooltip:=Perform Search").Click
'
'OutboundDeliverynumber = SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Maintain Outb\. Deliv\. Order - Warehouse No\. KZ01 \(Time Zone EST\)").SAPGuiGrid("micclass:=SAPGuiGrid","Index:=1").GetCellData(1,"Document")
'msgbox OutboundDeliverynumber
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/n/scwm/mon"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow").SAPGuiButton("name:=btn\[0\]").Click
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor").SAPGuiEdit("attachedtext:=Warehouse Number").Set "KZ01"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor").SAPGuiEdit("attachedtext:=Monitor").Set "SAP"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor").SAPGuiButton("tooltip:=Execute   \(F8\)").Click
'
'''SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiTree("micclass:=SAPGuiTree", "treetype:=SapListTree").Expand "Outbound;Documents;Outbound Delivery Order"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiTree("micclass:=SAPGuiTree", "treetype:=SapListTree").ActivateNode "Outbound;Documents;Outbound Delivery Order"
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","type:=GuiModalWindow","text:=/SCWM/SAPLWIP_DELIVERY_OUT").SAPGuiEdit("attachedtext:=Outb\. Delivery Order").Set OutboundDeliverynumber
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","type:=GuiModalWindow","text:=/SCWM/SAPLWIP_DELIVERY_OUT").SAPGuiButton("tooltip:=Execute   \(F8\)").Click
'''SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=/SCWM/SAPLWIP_DELIVERY_OUT").SAPGuiButton("tooltip:=Cancel   \(F12\)").Click
'wait(4)
'Set objToolbar = SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiToolbar("title:=Outbound Delivery Order")
'maxIndex = objToolbar.Object.ToolbarButtonCount 
'msgbox maxIndex
'strTooltip = "Wave"
'For i = 0 To maxIndex
'   If strTooltip = objToolbar.Object.GetToolbarButtonTooltip(i) then
'       strButtonID = objToolbar.Object.GetToolbarButtonId(i) 
'       msgbox strButtonID
'       objToolbar.Object.Presstoolbarbutton strButtonID
'       Exit For
'   End if
'Next @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf19.xml_;_
'
'WarehouseOrderNumber = SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiGrid("micclass:=SAPGuiGrid","title:=Wave").GetCellData(1,"Wave")
'msgbox WarehouseOrderNumber
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/n/scwm/rfui"
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiEdit("attachedtext:=Whse No\.").Set "KZ01"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiEdit("attachedtext:=Resource").Set "KMATHUR"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiEdit("attachedtext:=DefPresDvc").Set "Z002"
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiButton("tooltip:=Enter   \(Enter\)").Click
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiEdit("attachedtext:=WO No\.").Set WarehouseOrderNumber
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiButton("tooltip:=F4 Next").Click
'
'
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Dim mySendKeys,ObjSAPGuiTree,ObjKeyValues,arrNumbers(),ERPDeliverynumber,intCount, OutboundDeliverynumber, WarehouseOrderNumber @@ hightlight id_;_1442886_;_script infofile_;_ZIP::ssf2.xml_;_
'intCount=0
'SAPGuiUtil.AutoLogon "ECQ - SYK S4H QA", 900, "kmathur", "#1iskcon", "EN"
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=ECQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Purchase Requisition").SAPGuiOKCode("micclass:=SAPGuiOKCode").Set "/nME51N"
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=ECQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Purchase Requisition").SAPGuiButton("tooltip:=Enter").Click
'
'SAPGuiSession("micclass:=SAPGuiSession","systemname:=ECQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create Purchase Requisition").SAPGuiComboBox("name:=MEREQ_TOPLINE-BSART","type:=GuiComboBox").Select("Manual Req")



'SAPGuiSession("micclass:=SAPGuiSession","systemname:=EMQ").SAPGuiWindow("micclass:=SAPGuiWindow","text:=SAP").SAPGuiButton("tooltip:=F4 Next").Click
'SAPGuiSession("Session").SAPGuiWindow("Maintain Business Partner").SAPGuiTabStrip("GS_SCREEN_1200_TABSTRIP").Select "name := Tab_02"
'

'End If
'  If Instr(arrTabsName(i),"<Sales")>0Then 
'  msgbox arrTabsName(i)
'        SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign.Issue: Header Data").SAPGuiTabStrip("name:=TAXI_TABSTRIP_HEAD").Select arrTabsName(i)
'        Exit For
'    End If
'Next
''       msgbox strButtonID


'strTooltip = "Start immediately"
'For i = 0 To maxIndex
'   If strTooltip = objToolbar1.Object.GetToolbarButtonTooltip(i) then @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf18.xml_;_










'Set objGUiTabStrip = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign.Issue: Header Data").SAPGuiTabStrip("name:=TAXI_TABSTRIP_HEAD")
'strAllItems=SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign.Issue: Header Data").SAPGuiTabStrip("name:=TAXI_TABSTRIP_HEAD").GetROProperty("allitems") '//Just update the tabe name here
'arrTabsName=Split(strAllItems,";")' split and get tabs name in array
''iterate through all tabs
'Tabname = "Sales"
''For i=0 to Ubound(arrTabsName)
'For i=0 to Ubound(arrTabsName)
'msgbox arrTabsName(i)
'If arrTabsName(i) = Tabname Then
'	msgbox arrTabsName(i)
'        SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Create SYK Consign.Issue: Header Data").SAPGuiTabStrip("name:=TAXI_TABSTRIP_HEAD").Select arrTabsName(i)
'        Exit For
'    End If
'Next



'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","transaction:=/SCWM/PRDI").SAPGuiToolbar("name:=shell","Index:=1").ChildObjects


'SAPGuiSession("Session").SAPGuiWindow("Maintain Inbound Delivery").SAPGuiToolbar("ToolBarControl").PressButton "OK_OIP_TOGGLE" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf25.xml_;_
'maxIndex = objToolbar.ChildObjects
'	'msgbox maxIndex
'For i = 0 To maxIndex
''   		'If buttontext = objControl.Object.GetToolbarButtonTooltip(i) then
'       		msgbox maxIndex(i) 
'       		msgbox buttontext
'Next
'


'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","transaction:=/SCWM/PRDI").SAPGuiToolbar("id:=/app/con\[0\]/ses\[0\]/wnd\[0\]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell").PressContextButton "Goods Receipt + Save"
'msgbox maxIndex @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf22.xml_;_


'SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","transaction:=/SCWM/PRDI").SAPGuiButton("tooltip:=Perform Search").Click



'SAPGuiSession("Session").SAPGuiWindow("Maintain Inbound Delivery").Maximize @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf23.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Maintain Inbound Delivery").SAPGuiToolbar("ToolBarControl").PressContextButton "OIP_DETAIL_TO" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf23.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Maintain Inbound Delivery").SAPGuiToolbar("ToolBarControl").SelectMenuItem "Display Warehouse Tasks" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf24.xml_;_




'Set objToolbar = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","transaction:=/SCWM/PRDI").SAPGuiToolbar("id:=/app/con\[0\]/ses\[0\]/wnd\[0\]/usr/subSUB_COMPLETE_OIP:/SCWM/SAPLUI_DLV_PRD:2000/subSUB_OIP_DATA_AREA:/SCWM/SAPLUI_DLV_PRD:2210/cntlCONTAINER_TB_OIP_1/shellcont/shell")
'
''objToolbar.PressContextButton ("innertext:=OIP_DETAIL_TO")
' count = objToolbar.Object.ButtonCount
' msgbox count 
' For i = 0 To count-1
'tooltip =  objToolbar.Object.GetButtonTooltip(i)
'msgbox tooltip
'If tooltip ="Goods Receipt + Save" Then
'	buttonID = objToolbar.Object.GetButtonId(i)
'	objToolbar.PressButton buttonID
'	Exit For
'End If
'msgbox buttonicon
'
''objToolbar.PressButton("OIP_DETAIL_TO")
''msgbox maxIndex.count
'
'Next

'val = "Comma,Separated"
'arrData = Split(val, ",")
'count = ubound(arrData)+1
'msgbox count
'arrData(1) = "hello"
'For i = 0 To ubound(arrData) 
'	msgbox arrData(i)
'Next
'	
'	Set objControl = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Stock Overview: Basic List").SAPGuiTree("micclass:=SAPGuiTree")
'	Set ObjSAPGuiTree = objControl.Object
'
'    'Get the node keys , a key is a number/position in the 
'    'A key value starts from 1.
'    Set ObjKeyValues = ObjSAPGuiTree.GetAllNodeKeys
'    'get the total count
'    ''This count indicates the number of items/nodes in the Tree
'    intNodeCount = ObjKeyValues.Count 
'    
'    Set colHeaders = ObjSAPGuiTree.GetColumnHeaders
''    
'    intcolHeadersCount = colHeaders.count
'	msgbox intcolHeadersCount
'    'Iterate through the nodes of the tree
'    For i = 0 to intNodeCount-1
'        'Get the node text
'        strNodeText=ObjSAPGuiTree.GetNodeTextByKey(ObjKeyValues(i))
'        
'        If Instr(strNodeText,strValue)>0 Then
''            msgbox strNodeText
''            docOverviewDate = ObjSAPGuiTree.GetItemText(ObjKeyValues(i),colHeaders(1))
''            msgbox docOverviewDate
''            docOverviewStatus = ObjSAPGuiTree.GetItemText(ObjKeyValues(i),colHeaders(2))
''            msgbox docOverviewStatus
'
''			docOverviewDate = ObjSAPGuiTree.GetItemText(ObjKeyValues(i),colHeaders(1))
'			
'			For j = 0 To intcolHeadersCount-1
'				key = ObjSAPGuiTree.GetItemText(ObjKeyValues(0),colHeaders(j))
'				msgbox key
'			Next 
'			
'            Exit For
'        End if
'    Next

'buttontext = "More methods"
'
'	Set objControl = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiToolbar("title:=Warehouse Task")
'	maxIndex = objControl.Object.ToolbarButtonCount 
'	
'msgbox maxIndex
'	'tooltip =  objToolbar.Object.GetButtonTooltip(i)
'	For i = 0 To maxIndex-1
'	If buttontext = objControl.Object.GetToolbarButtonTooltip(i) then
'       		strButtonID = objControl.Object.GetToolbarButtonId(i) 
'   		
'			objControl.PressButton strButtonID
'       		Exit For
'   		End if
'	Next
'	
	
	
''buttontext = "More methods"
'Set objControl = SAPGuiSession("micclass:=SAPGuiSession").SAPGuiWindow("micclass:=SAPGuiWindow","text:=Warehouse Management Monitor SAP - Warehouse Number KZ01").SAPGuiGrid("title:=Warehouse Task")
''maxIndex = objControl.Object.ToolbarButtonCount 
'child = objControl.Object.Children 
'count = child.Count
''	msgbox maxIndex
'	For i = 0 To count-1
'	
'	msgbox child(i).Name
'	Next
'	Set child = Nothing
'   		'If buttontext = objControl.Object.GetToolbarButtonTooltip(i) then
''       		strButtonID = objControl.Object.GetToolbarButtonId(i) 
''       		'msgbox objControl.Object.GetToolbarButtonTooltip(i)
''       		'msgbox strButtonID
''       		objControl.Object.PressContexttoolbarbutton strButtonID
'' @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf30.xml_;_
''      		Exit For
''   		End if
''	Next
	
	
 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf27.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Warehouse Management Monitor").SAPGuiToolbar("GridToolbar_2").PressContextButton "METHODS" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf28.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Warehouse Management Monitor").SAPGuiToolbar("GridToolbar_2").SelectMenuItemById "@M00008" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf29.xml_;_


'SAPGuiSession("Session").SAPGuiWindow("Warehouse Management Monitor").SAPGuiToolbar("GridToolbar").PressButton "METHODS" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf31.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Warehouse Management Monitor").SAPGuiToolbar("GridToolbar").PressContextButton "METHODS" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf32.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Warehouse Management Monitor").SAPGuiToolbar("GridToolbar").SelectMenuItemById "@M00008" @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf33.xml_;_
