# AutoHotkey-Midi

Add MIDI input event handling to your AutoHotkey scripts

```ahk

#include AutoHotkey-Midi/Midi.ahk

midi := new Midi()
midi.OpenMidiOutByName("X-TOUCH MINI")
midi.OpenMidiInByName("X-TOUCH MINI")

; send some  Outout
midi.MidiOut("CC", 1, 127, 0) ; ControllerChange on Channel 1, Code 27
midi.MidiOut("N1", 1, 1, 100) ; Note on On Channel 1, Note 1, Velocity 100

Return

MidiNoteOnA4:
	MsgBox You played note A4!
	Return

MidiControlChange1:
	cc := midi.MidiIn()
	ccValue := cc.value
	MsgBox You set the mod wheel to %ccValue%
	Return

```

## Requirements

* A modern version of AutoHotKey (1.1+) from http://ahkscript.org/
* A system with winmm.dll (Windows 2000 or greater)

## License

BSD

## TODO

* Documentation!
