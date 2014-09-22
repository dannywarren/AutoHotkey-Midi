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

; Defines the size of the standard chromatic scale
Global MIDI_NOTE_SIZE := 12

; Defines the midi notes 
Global MIDI_NOTES     := [ "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B" ]

; Defines the octaves for midi notes
Global MIDI_OCTAVES   := [ -2, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8 ]


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
Global midiLabel := "Midi"

; Enable or disable label event handling
Global midiLabelEvents  := True

; Enable or disable lazy midi in event debugging via tooltips
Global midiEventTooltips := False


; Midi class interface
Class Midi
{

  midiInDevice := -1

  ; Instance creation
  __New( newMidiInDevice:=-1 )
  {

    this.SetMidiInDevice( newMidiInDevice )
    this.LoadMidi()

  }


  ; Instance destruction
  __Delete()
  {

    this.StopMidiIn()
    this.UnloadMidi()

  }


  ; Set the current midi in device
  SetMidiInDevice( newMidiInDevice )
  {

    ; Bail out if this is already the midi device we are using
    If ( newMidiInDevice == this.midiInDevice )
      Return
 
    ; Bail out if no new midi in device was given
    If ( newMidiInDevice < 0 )
      Return

    ; Stop listening to the current midi device (if applicable)
    this.StopMidiIn()

    ; Set the new midi in device and then start listening to it
    this.midiInDevice := newMidiInDevice
    this.StartMidiIn()

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

  ; Will hold the labels we call so the user can capture this midi event, we
  ; always start with a generic ":Midi" label so it always gets called first
  labelCallbacks := [ midiLabel ]

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

  ; Add a label callback for the status, ie ":MidiNoteOn"
  labelCallbacks.Insert( midiLabel . midiEvent.status )

  ; Determine how to handle the one or two data bytes sent along with the event
  ; based on what type of status event was seen
  if ( midiEvent.status == "NoteOff" || midiEvent.status == "NoteOn" || midiEvent.status == "AfterTouch" )
  {

    ; Store the raw note number and velocity data
    midiEvent.noteNumber  := data1
    midiEvent.velocity    := data2

    ; Figure out which chromatic note this note number represents
    noteScaleNumber := Mod( midiEvent.noteNumber, MIDI_NOTE_SIZE )

    ; Look up the name of the note in the scale
    midiEvent.note := MIDI_NOTES[ noteScaleNumber + 1 ]

    ; Determine the octave of the note in the scale 
    noteOctaveNumber := Floor( midiEvent.noteNumber / MIDI_NOTE_SIZE )

    ; Look up the octave for the note
    midiEvent.octave := MIDI_OCTAVES[ noteOctaveNumber + 1 ]

    ; Create a friendly name for the note and octave, ie: "C4"
    midiEvent.noteName := midiEvent.note . midiEvent.octave

    ; Add label callbacks for notes, ie ":MidiNoteOnA", ":MidiNoteOnA5", ":MidiNoteOn97"
    labelCallbacks.Insert( midiLabel . midiEvent.status . midiEvent.note )
    labelCallbacks.Insert( midiLabel . midiEvent.status . midiEvent.noteName )
    labelCallbacks.Insert( midiLabel . midiEvent.status . midiEvent.noteNumber )

  }
  else if ( midiEvent.status == "ControlChange" )
  {

    ; Store controller number and value change
    midiEvent.controller := data1
    midiEvent.value      := data2

    ; Add label callback for this controller change, ie ":MidiControlChange12"
    labelCallbacks.Insert( midiLabel . midiEvent.status . midiEvent.controller )

  }
  else if ( midiEvent.status == "ProgramChange" )
  {

    ; Store program number change
    midiEvent.program := data1

    ; Add label callback for this program change, ie ":MidiProgramChange2"
    labelCallbacks.Insert( midiLabel . midiEvent.status . midiEvent.program )

  }
  else if ( midiEvent.status == "ChannelPressure" )
  {
    
    ; Store pressure change value
    midiEvent.pressure := data1

  }
  else if ( midiEvent.status == "PitchWheel" )
  {

    ; Store pitchwheel change, which is a combination of both data bytes 
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
    
    ; Add label callback for sysex event, ie: ":MidiClock" or ":MidiStop"
    labelCallbacks.Insert( midiLabel . midiEvent.sysex )

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

  ; Iterate over all the label callbacks we built during this event and jump
  ; to them now (if they exist elsewhere in the code)
  If ( midiLabelEvents )
  {
    For labelIndex, labelName In labelCallbacks
    {
      If IsLabel( labelName )
        Gosub %labelName%
    }   
  }

  ; Call debugging if enabled
  __MidiInEventDebug( midiEvent )

}


; Send event information to a listening debugger
__MidiInEventDebug( midiEvent )
{

  debugStr := ""

  For key, value In midiEvent
    debugStr .= key . ":" . value . "`n"

  debugStr .= "---`n"

  ; Always output event debug to any listening debugger
  OutputDebug, % debugStr 

  ; If lazy tooltip debugging is enabled, do that too
  if midiEventTooltips
    ToolTip, % debugStr

}

