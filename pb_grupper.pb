OpenWindow(1, 0, 0, 200, 200, "")

Global grupper = ButtonGadget(#PB_Any,20,5,100,20,"Gruppfiler")

Repeat
  
  event = WindowEvent()
  
  If event = #PB_Event_Gadget
    nr = EventGadget()

    If nr = grupper
      file$ = OpenFileRequester("Öppna", "","",0)
      a$ = "klock_grupper.py " + file$
      RunProgram("python",a$,"")
      
    EndIf
    
  EndIf
  
  
  
  Delay(1)
  
Until event = #PB_Event_CloseWindow

; IDE Options = PureBasic 5.71 LTS (Windows - x64)
; CursorPosition = 13
; EnableXP
; Executable = analysplotter.exe