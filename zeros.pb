Enumeration
  #WINDOW_MAIN
  #TEXT_ID
  #STRING_ID
  #TEXT_NAME
  #STRING_NAME
  #BTN_CANCEL
  #BTN_USE
EndEnumeration

Global Quit.b = #False

#FLAGS = #PB_Window_ScreenCentered | #PB_Window_SystemMenu

If OpenWindow(#WINDOW_MAIN, 0, 0, 480, 150, "Leading Zeroes", #FLAGS)
  If UseGadgetList(WindowID(#WINDOW_MAIN))
    TextGadget(#TEXT_ID, 20, 22, 120, 18, "Spreadsheet ID:", #PB_Text_Right)
    StringGadget(#STRING_ID, 140, 20, 330, 20, GetClipboardText())
    
    TextGadget(#TEXT_NAME, 20, 52, 120, 18, "Sheet name:" ,#PB_Text_Right)
    StringGadget(#STRING_NAME, 140, 50, 330, 20, "Export.xls")
    
    ButtonGadget(#BTN_CANCEL, 20, 100, 100, 25, "Cancel", #PB_Button_Default)
    ButtonGadget(#BTN_USE, 330, 100, 100, 25, "Use", #PB_Button_Default)
    
    SetActiveGadget(#BTN_CANCEL)
  
    Repeat
      Event = WaitWindowEvent()
      Select Event
        Case #PB_Event_Gadget
          Select EventGadget()
            Case #BTN_CANCEL
              Quit = #True
            Case #BTN_USE
              If RunProgram("python3","leading_zero.py --id="+GetGadgetText(#STRING_ID)+" --sheetname="+GetGadgetText(#STRING_NAME), GetCurrentDirectory())
                ;MessageRequester("Leading zero script", "Leading Zero executed successfully", #PB_MessageRequester_Ok | #PB_MessageRequester_Info)
                Quit = #True
              EndIf 
          EndSelect
      EndSelect
    Until Event = #PB_Event_CloseWindow Or Quit
  EndIf 
EndIf

End
; IDE Options = PureBasic 5.72 (Linux - x64)
; CursorPosition = 38
; EnableXP
; Executable = zeros