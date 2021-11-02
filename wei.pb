; Windows Experience Index
; Alex, 2021/11/02

; Configuration
Global directory$ = "C:\Windows\Performance\WinSAT\DataStore\"
Global pattern$   = "*Formal*Recent*.xml"
Global xmlpath$   = "/WinSAT/WinSPR/*"
Global cmd$       = "C:\Windows\sysnative\cmd.exe" ; Override 64-bit-redirect
Global cmdParam$  = "/c winsat formal" ; Override 64-bit-redirect

; Variables
Structure performanceStructure 
    value.s
    line.i
EndStructure
Global NewMap performance.performanceStructure()


; Procedures
Declare.s GetLatestFile()
Declare Parse(file$)
Declare Populate()
Declare ResetValues()
Declare Redraw(Gadget)

; GUI
XIncludeFile "wei-window.pbf"
OpenWindowMain()
Global subheadline$ = GetGadgetText(Subheadline)

; ------------------------------------------------------------------------------------------------

; Init
ResetValues()
Parse(GetLatestFile())
Populate()

; Window
Repeat
  Event = WaitWindowEvent()
  
  If Event = #PB_Event_Gadget And EventGadget() = Button
    ; Reset values
    ResetValues()
    Populate()
    
    ; Recalculate
    oldValue$ = GetGadgetText(Button)
    SetGadgetText(Button, "Calculating ...")
    If Not RunProgram(cmd$, cmdParam$, "", #PB_Program_Wait | #PB_Program_Hide)
      MessageRequester("Error", "Unable To run WinSAT");
    EndIf
    SetGadgetText(Button, oldValue$)
    
    ; Load new values
    Parse(GetLatestFile())
    Populate()
  EndIf
  
Until Event = #PB_Event_CloseWindow

; ------------------------------------------------------------------------------------------------

; Populate window
Procedure Populate()

  ; Can not access in ForEach-loop
  SystemScore$ = performance("SystemScore")\value
  ; List
  ForEach performance()
    Debug MapKey(performance()) + " = " + performance()\value + " - " + performance()\line
    If performance()\line >= 0
      SetGadgetItemText(List, performance()\line, performance()\value, 2)
      ; Mark lowest value
      If performance()\value = SystemScore$ And performance()\value <> "-"
        SetGadgetItemColor(List, performance()\line, #PB_Gadget_BackColor, GetSysColor_(#COLOR_BTNFACE), 2)
      EndIf
    EndIf
  Next

  ; Base score
  SetGadgetText(Score,performance("SystemScore")\value)
  SetGadgetText(Subheadline,subheadline$ + " " + performance("SystemScore")\value)
  
  Redraw(List)
EndProcedure

; Parse
Procedure Parse(file$)
  ; Variables
  Enumeration
    #XML
  EndEnumeration
  
  ; Load
  If file$ = "" Or LoadXML(#XML, directory$ + file$) = 0
    MessageRequester("Error", "Unable to load XML " + file$)
    ProcedureReturn
  EndIf
  
  ; Check
  If XMLStatus(#XML) <> #PB_XML_Success
    Message$ = "Error in the XML file:" + Chr(13)
    Message$ + "Message: " + XMLError(#XML) + Chr(13)
    Message$ + "Line: " + Str(XMLErrorLine(#XML)) + "   Character: " + Str(XMLErrorPosition(#XML))
    MessageRequester("Error", Message$)
    ProcedureReturn
  EndIf
  
  ; Parse
  *node = XMLNodeFromPath(MainXMLNode(#XML),xmlpath$)
  While *node <> 0
    ; Debug "- " + GetXMLNodeName(*node) + " = " + GetXMLNodeText(*node)
    ; Store only specific values
    If FindMapElement(performance(), GetXMLNodeName(*node))
      performance(GetXMLNodeName(*node))\value = GetXMLNodeText(*node)
    EndIf
    *node = NextXMLNode(*node)
  Wend  
EndProcedure

; Get latest file
Procedure.s GetLatestFile()
  xml$ = ""
  start = 0
  If ExamineDirectory(0, directory$, pattern$)  
    While NextDirectoryEntry(0)
      ; Check for date
      If DirectoryEntryDate(0, #PB_Date_Created) > start
        xml$ = DirectoryEntryName(0)
        start = DirectoryEntryDate(0, #PB_Date_Created)
      EndIf
    Wend
    FinishDirectory(0)
  Else
    MessageRequester("Error", "Unable to open directory " + directory$)
  EndIf
  ProcedureReturn xml$
EndProcedure

Procedure ResetValues()
  performance("SystemScore")\value   = "-"
  performance("SystemScore")\line   = -1
  performance("CpuScore")\value      = "-"
  performance("CpuScore")\line       = 0
  performance("MemoryScore")\value   = "-"
  performance("MemoryScore")\line    = 1
  performance("GraphicsScore")\value = "-"
  performance("GraphicsScore")\line  = 2
  performance("GamingScore")\value   = "-"
  performance("GamingScore")\line    = 3
  performance("DiskScore")\value     = "-"
  performance("DiskScore")\line      = 4
EndProcedure

Procedure Redraw(Gadget)
  SendMessage_(GadgetID(Gadget),#WM_SETREDRAW,1,0)
  InvalidateRect_(GadgetID(Gadget),0,0)
  UpdateWindow_(GadgetID(Gadget))
EndProcedure

; IDE Options = PureBasic 5.73 LTS (Windows - x86)
; CursorPosition = 24
; Folding = -
; EnableXP
; UseIcon = icon.ico
; Executable = wei.exe