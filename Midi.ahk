;
; Midi.ahk
; Add MIDI input event handling to your AutoHotkey scripts
;
; Danny Warren <danny@dannywarren.com>
; https://github.com/dannywarren/AutoHotkey-Midi
;


; Always use gui mode when using the midi library, since we need something to
; attach midi events to
Gui, +LastFound


; Defines for midi event callbacks (see mmsystem.h)
Global MIDI_CALLBACK_WINDOW   := 0x10000
Global MIDI_CALLBACK_TASK     := 0x20000
Global MIDI_CALLBACK_FUNCTION := 0x30000

; Defines for midi event types (see mmsystem.h)
Global MIDI_OPEN      := 0x3C1
Global MIDI_CLOSE     := 0x3C2
Global MIDI_DATA      := 0x3C3
Global MIDI_LONGDATA  := 0x3C4
Global MIDI_ERROR     := 0x3C5
Global MIDI_LONGERROR := 0x3C6
Global MIDI_MOREDATA  := 0x3CC


; This is where we will keep the most recent midi in event data so that it can
; be accessed via the Midi object, since we cannot store it in the object due
; to how events work
; We will store the last event by the handle used to open the midi device, so
; at least we won't clobber midi events from other instances of the object
Global __midiInEvent := {}

; Track the number of instances that may be listening, so that we can know when
; to set up and tearn down event callbacks
Global __midiInListeners := 0

; Default label names
Global __midiLabel   := "Midi"
Global __midiInLabel := "MidiIn"

; Enable or disable label event handling
Global __midiInLabelEvents  := True

; Enable or disable lazy midi in event debugging
Global __midiInDebug := False


; Midi class interface
Class Midi
{

  midiInDevice := -1

  ; Instance creation
  __New( newMidiInDevice )
  {

    ; Until we implement a better setter/getter for the midi device and
    ; add it to the menu options, we will just pass the id in 
    if ( newMidiInDevice < 0 )
    {
      MsgBox, No midi input device specified
      ExitApp
    }

    this.midiInDevice := newMidiInDevice

    this.LoadMidi()
    this.StartMidiIn()

  }


  ; Instance destruction
  __Delete()
  {

    this.StopMidiIn()
    this.UnloadMidi()

  }


  ; Load midi dlls
  LoadMidi()
  {
    
    this.midiDll := DllCall( "LoadLibrary", "Str", "winmm.dll", "Ptr" )
    
    If ( ! this.midiDll )
    {
      MsgBox, Missing system midi library winmm.dll
      ExitApp
    }

  }


  ; Unload midi dlls
  UnloadMidi()
  {

    If ( this.midiDll )
    {
      DllCall( "FreeLibrary", "Ptr", this.midiDll )
    }

  }


  ; Open midi in device and start listening
  StartMidiIn()
  {

    ; Bail out if we don"t have a midi in device defined
    If ( ! this.midiInDevice )
    {
      Return
    }

    ; Create variable to store the handle the dll open will give us
    ; NOTE: Creating variables this way doesn't work with class variables, so
    ; we have to create it locally and then store it in the class later after
    VarSetCapacity( midiIn, 4, 0 )

    ; Get the autohotkey window to attach the midi events to
    thisWindow := WinExist()

    ; Open the midi device and attach event callbacks
    midiInOpenResult := DllCall( "winmm.dll\midiInOpen", UINT, &midiIn, UINT, this.midiInDevice, UINT, thisWindow, UINT, 0, UINT, MIDI_CALLBACK_WINDOW )

    ; Error handling
    If ( midiInOpenResult || ! midiIn )
    {
      MsgBox, Failed to open midi in device
      ExitApp
    }

    ; Fetch the actual handle value from the pointer
    midiIn := NumGet( midiIn, UINT )

    ; Start monitoring midi signals
    midiInStartResult := DllCall( "winmm.dll\midiInStart", UINT, midiIn )

    ; Error handling
    If ( midiInStartResult )
    {
      MsgBox, Failed to start midi in device
      ExitApp
    }

    ; Save the midi input handle to our class instance now that we are done
    ; messing with it
    this.midiIn := midiIn

    ; Create a spot in our global event storage for this midi input handle
    __MidiInEvent[this.midiIn] := {}

    ; Register a callback for each midi event
    ; We only need to do this once per instance of our class, so if another
    ; instance already did it then we don't need to
    if ( ! __midiInListeners )
    {

      OnMessage( MIDI_OPEN,      "__MidiInCallback" )
      OnMessage( MIDI_CLOSE,     "__MidiInCallback" )
      OnMessage( MIDI_DATA,      "__MidiInCallback" )
      OnMessage( MIDI_LONGDATA,  "__MidiInCallback" )
      OnMessage( MIDI_ERROR,     "__MidiInCallback" )
      OnMessage( MIDI_LONGERROR, "__MidiInCallback" )
      OnMessage( MIDI_MOREDATA,  "__MidiInCallback" )
      
    }

    ; Increment the number of midi listeners
    __midiInListeners++

  }


  ; Close midi in device and stop listening
  StopMidiIn()
  {

    ; Bail out if we don"t have a midi in handle open
    If ( ! this.midiIn )
    {
      Return
    }

    ; Unregister callbacks if we are the last 
    if ( __midiInListeners <= 1 )
    {
      OnMessage( MIDI_OPEN,      "" )
      OnMessage( MIDI_CLOSE,     "" )
      OnMessage( MIDI_DATA,      "" )
      OnMessage( MIDI_LONGDATA,  "" )
      OnMessage( MIDI_ERROR,     "" )
      OnMessage( MIDI_LONGERROR, "" )
      OnMessage( MIDI_MOREDATA,  "" )
    }

    ; Decrement the number of midi listeners
    __midiInListeners--

    ; Destroy any midi in events that might be left over
    __MidiInEvent[this.midiIn] := {}

    ; Stop monitoring midi
    midiInStopResult = DllCall( "winmm.dll\midiInStop", UINT, this.midiIn )

    ; Error handling
    If ( midiInStartResult )
    {
      MsgBox, Failed to stop midi in device
      ExitApp
    }

    ; Close the midi handle
    midiInStopResult = DllCall( "winmm.dll\midiInClose", UINT, this.midiIn )

    ; Error handling
    If ( midiInStartResult )
    {
      MsgBox, Failed to close midi in device
      ExitApp
    }

  }


  ; Returns the last midi in event
  MidiInEvent()
  {

    If ( ! this.midiIn )
    {
      Return
    }

    Return __MidiInEvent[this.midiIn]

  }


}


; Event callback for midi input event
; Note that since this is a callback method, it has no concept of "this" and
; can't access class members
__MidiInCallback( wParam, lParam, msg )
{

  ; Will hold the midi event object we are building for this event
  midiEvent := {}

  ; Grab the raw midi bytes
  rawBytes := lParam

  ; Split up the raw midi bytes as per the midi spec
  highByte  := lParam & 0xF0 
  lowByte   := lParam & 0x0F
  data1     := (lParam >> 8) & 0xFF
  data2     := (lParam >> 16) & 0xFF

  ; Determine the friendly name of the midi event based on the status byte
  if ( highByte == 0x80 )
  {
    midiEvent.status := "NoteOff"
  }
  else if ( highByte == 0x90 )
  {
    midiEvent.status := "NoteOn"
  }
  else if ( highByte == 0xA0 )
  {
    midiEvent.status := "Aftertouch"
  }
  else if ( highByte == 0xB0 )
  {
    midiEvent.status := "ControlChange"
  }
  else if ( highByte == 0xC0 )
  {
    midiEvent.status := "ProgramChange"
  }
  else if ( highByte == 0xD0 )
  {
    midiEvent.status := "ChannelPressure"
  }
  else if ( highByte == 0xE0 )
  {
    midiEvent.status := "PitchWheel"
  }
  else if ( highByte == 0xF0 )
  {
    midiEvent.status := "Sysex"
  }
  else
  {
    Return
  }

  ; Determine how to handle the one or two data bytes sent along with the event
  ; based on what type of status event was seen
  if ( midiEvent.status == "NoteOff" || midiEvent.status == "NoteOn" || midiEvent.status == "AfterTouch" )
  {
    midiEvent.note     := data1
    midiEvent.velocity := data2
  }
  else if ( midiEvent.status == "ControlChange" )
  {
    midiEvent.number := data1
    midiEvent.value  := data2
  }
  else if ( midiEvent.status == "ProgramChange" )
  {
    midiEvent.program := data1
  }
  else if ( midiEvent.status == "ChannelPressure" )
  {
    midiEvent.pressure := data1
  }
  else if ( midiEvent.status == "PitchWheel" )
  {
    midiEvent.pitch := ( data2 << 7 ) + data1
  }
  else if ( midiEvent.status == "Sysex" )
  {

    ; Sysex messages have another status byte that indicates which type of sysex
    ; message it is (the high byte, which is normally used for the midi channel,
    ; is used for this instead)
    if ( lowByte == 0x0 )
    {
      midiEvent.sysex := "SysexData"
      midiEvent.data  := byte1
    }
    if ( lowByte == 0x1 )
    {
      midiEvent.sysex := "Timecode"
    }
    if ( lowByte == 0x2 )
    {
      midiEvent.sysex     := "SongPositionPointer"
      midiEvent.position  := ( data2 << 7 ) + data1
    }
    if ( lowByte == 0x3 )
    {
      midiEvent.sysex   := "SongSelect"
      midiEvent.number  := data1
    }
    if ( lowByte == 0x6 )
    {
      midiEvent.sysex := "TuneRequest"
    }
    if ( lowByte == 0x8 )
    {
      midiEvent.sysex := "Clock"
    }
    if ( lowByte == 0x9 )
    {
      midiEvent.sysex := "Tick"
    }
    if ( lowByte == 0xA )
    {
      midiEvent.sysex := "Start"
    }
    if ( lowByte == 0xB )
    {
      midiEvent.sysex := "Continue"
    }
    if ( lowByte == 0xC )
    {
      midiEvent.sysex := "Stop"
    }
    if ( lowByte == 0xE )
    {
      midiEvent.sysex := "ActiveSense"
    }
    if ( lowByte == 0xF )
    {
      midiEvent.sysex := "Reset"
    }
    
  }

  ; Channel is always handled the same way for all midi events except sysex
  if ( midiEvent.status != "Sysex" )
  {
    midiEvent.channel := lowByte + 1
  }

  ; Always include the raw midi data, just in case someone wants it
  midiEvent.rawBytes  := rawBytes
  midiEvent.highByte  := highByte
  midiEvent.lowByte   := lowByte
  midiEvent.data1     := data1
  midiEvent.data2     := data2

  ; Store this midi in event in our global array of midi messages, so that the
  ; appropriate midi class an access it later
  __MidiInEvent[wParam] := midiEvent

  ; Call even labels for this event if label handling is enabled
  if ( __midiInLabelEvents )
  {
    __MidiInLabel( midiEvent )
  }

  ; Call debugging if enabled
  if ( __midiInDebug )
  {
    __MidiInDebug( midiEvent )
  }

}


; For a given midi event, call any global labels that are applicable to the
; midi event
__MidiInLabel( midiEvent )
{

  ; Always call the generic midi event label, ie: ":Midi"
  midiLabel := __midiLabel
  If IsLabel( midiLabel )
    Gosub %midiLabel%

  ; Always call the generic midi in event label, ie ":MidiIn"
  midiInLabel := __midiInLabel
  If IsLabel( midiInLabel )
    Gosub %midiInLabel%

  ; Call a label for the specific midi status, ie ":MidiNoteOn"
  midiStatusLabel := midiLabel . midiEvent.status
  If IsLabel( midiStatusLabel )
    Gosub %midiStatusLabel%

  ; Call a label for the specific midi in status, ie ":MidiInNoteOn"
  midiInStatusLabel := midiInLabel . midiEvent.status
  If IsLabel( midiInStatusLabel )
    Gosub %midiInStatusLabel%

  ; Add labels for sysex message sub-statuses if applicable
  if ( midiEvent.status == "Sysex" )
  {

    ; Call a label for the specific sysex event, ie ":MidiClock"
    midiSysexLabel := midiLabel . midiEvent.sysex 
    If IsLabel( midiSysexLabel )
      Gosub %midiSysexLabel%

    ; Call a label for the specific sysex event, ie ":MidiInClock"
    midiInSysexLabel := midiInLabel . midiEvent.sysex 
    If IsLabel( midiInSysexLabel )
      Gosub %midiInSysexLabel%

  }

}


; Tooltip containing all the data from a midi event
__MidiInDebug( midiEvent )
{

  debugStr := ""

  for key, value in midiEvent
    debugStr .= key . ":" . value . "`n"

  ToolTip, %debugStr%

}

