# AutoHotkey-Midi

AutoHotkey-Midi adds MIDI input functionality to AutoHotkey.

```ahk

#include AutoHotkey-Midi/Midi.ahk

midi := new Midi( 5 )

MidiNoteOn:
	
	midiEvent := midi.MidiInEvent()

	if ( midiEvent.note == 57 ) 
	{
		MsgBox You played A4!
	}
	
	Return

```

## Requirements

* A modern version of AutoHotKey (1.1+) from http://ahkscript.org/
* A system with winmm.dll (Windows 2000 or greater)

## License

BSD

## TODO

* Midi output event support
* Midi device selection in autohotkey menu
* Midi reinit on device change
* More label jumps based on data and not just event status (like "MidiNoteOnA5" etc)
* Object scoped callbacks (most likely not possible with pure autothotkey)
