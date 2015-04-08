import win32api
import time
import random
import argparse
import sys
import os
import argparse
from mido import MidiFile
import ctypes

#http://stackoverflow.com/questions/14489013/simulate-python-keypresses-for-controlling-a-game
SendInput = ctypes.windll.user32.SendInput

# C struct redefinitions 
PUL = ctypes.POINTER(ctypes.c_ulong)
class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]

class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time",ctypes.c_ulong),
                ("dwExtraInfo", PUL)]

class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]

# Actuals Functions
def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra) )
    x = Input( ctypes.c_ulong(1), ii_ )
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

# directx scan codes http://www.gamespp.com/directx/directInputKeyboardScanCodes.html
def KeyPress(k):
    PressKey(k) 
    #time.sleep(.005)
    ReleaseKey(k)

directInputKeyboardScanCodes = {'1': 0x02, 
                               '2': 0x03, 
                               '3': 0x04, 
                               '4': 0x05, 
                               '5': 0x06, 
                               '6': 0x07, 
                               '7': 0x08, 
                               '8': 0x09, 
                               '9': 0x0A, 
                               '0': 0x0B,

                               'q': 0x10, 
                               'w': 0x11, 
                               'e': 0x12, 
                               'r': 0x13, 
                               't': 0x14, 
                               'y': 0x15, 
                               'u': 0x16, 
                               'i': 0x17, 
                               'o': 0x18, 
                               'p': 0x19,

                               'a': 0x1E, 
                               's': 0x1F, 
                               'd': 0x20, 
                               'f': 0x21, 
                               'g': 0x22, 
                               'h': 0x23, 
                               'j': 0x24, 
                               'k': 0x25, 
                               'l': 0x26,

                               'z': 0x2c, 
                               'x': 0x2d, 
                               'c': 0x2e, 
                               'v': 0x2f, 
                               'b': 0x30, 
                               'n': 0x31, 
                               'm': 0x32 }

keybinds = []

def SendToGame(str):
    global directInputKeyboardScanCodes

    for c in str:
        KeyPress(directInputKeyboardScanCodes[c])
        #time.sleep(0.1)
    #time.sleep(0.1)


prev_octave = 1

def NewNote(note, velo):
    global keybinds
    global prev_octave

    #
    # note 60 (middle C) (C4)
    #
    # C   C#  D   D#  E   F   F#  G   G#  A   A#  B 
    # 48  49  50  51  52  53  54  55  56  57  58  59
    # 60  61  62  63  64  65  66  67  68  69  70  71
    # 72  73  74  75  76  77  78  79  80  81  82  83
    if (note < 48 or note > 83):
        return
    
    print("New note:", note)

    octave = 1
    if (note >= 48 and note < 60):
        octave = 0
    if (note >= 60 and note < 72):
        octave = 1
    if (note >= 72 and note <= 84):
        octave = 2

    print("octave:", octave, "  prev_octave:", prev_octave)
    octave_save = octave

    # bindings: 
    key_prev_oct = keybinds[8]
    key_next_oct = keybinds[9]

    keystring = ""

    #print("oct:", octave)
    if (octave > prev_octave):
        octave -= 1
        keystring = keystring + key_next_oct
        #print("oct:", octave)
    if (octave > prev_octave):
        octave -= 1
        keystring = keystring + key_next_oct
        #print("oct:", octave)

    if (octave < prev_octave):
        octave += 1
        keystring = keystring + key_prev_oct
    if (octave < prev_octave):
        octave += 1
        keystring = keystring + key_prev_oct

    #
    # note 60 (middle C) (C4)
    #
    # C   C#  D   D#  E   F   F#  G   G#  A   A#  B 
    # 48  49  50  51  52  53  54  55  56  57  58  59
    # 60  61  62  63  64  65  66  67  68  69  70  71
    # 72  73  74  75  76  77  78  79  80  81  82  83
    if (note % 12 == 0 or note % 12 == 1):
        keystring = keystring + keybinds[0]
    if (note % 12 == 2 or note % 12 == 3):
        keystring = keystring + keybinds[1]
    if (note % 12 == 4):
        keystring = keystring + keybinds[2]
    if (note % 12 == 5 or note % 12 == 6):
        keystring = keystring + keybinds[3]
    if (note % 12 == 7 or note % 12 == 8):
        keystring = keystring + keybinds[4]
    if (note % 12 == 9 or note % 12 == 10):
        keystring = keystring + keybinds[5]
    if (note % 12 == 11):
        keystring = keystring + keybinds[6]
    if (note == 84):
        keystring = keystring + keybinds[7]

    print("str:", keystring);

    prev_octave = octave_save

    SendToGame(keystring)


def PlayAllFile(mid, speed):
    for message in mid:
        if (message.type == 'note_on'):
            sleept1 = message.time / speed;
            if (sleept1 > 60):
                sleept1 = 60
            time.sleep(sleept1)
            print(message)
            NewNote(message.note, message.velocity)

        if (message.type == 'note_off'):
            sleept1 = message.time / speed;
            if (sleept1 > 60):
                sleept1 = 60
            time.sleep(sleept1)
            print(message)


def PlayTrack(track, speed, ticks_per_beat):

    tempo = 500000
    seconds_per_tick = 0.0005

    for message in track:

        if (message.type == 'set_tempo'):
            tempo = message.tempo
            print(message)
            seconds_per_beat = tempo / 1000000.0
            seconds_per_tick = seconds_per_beat / float(ticks_per_beat)
            print("seconds_per_beat", seconds_per_beat, "  seconds_per_tick", seconds_per_tick)

        if (message.type == 'note_on'):
            sleept1 = seconds_per_tick * message.time / speed;
            if (sleept1 > 60):
                sleept1 = 60
            time.sleep(sleept1)
            print(message)
            NewNote(message.note, message.velocity)

        if (message.type == 'note_off'):
            sleept1 = seconds_per_tick * message.time / speed;
            if (sleept1 > 60):
                sleept1 = 60
            time.sleep(sleept1)
            print(message)

def PlaySync(sync_count):
    for i in range(0,sync_count):
        NewNote(60, 60)
        time.sleep(0.25)
    time.sleep(1)

def ResetInstrument():
    global keybinds
    print("Setting initial octave.")
    SendToGame(keybinds[8])
    time.sleep(0.2)
    SendToGame(keybinds[8])
    time.sleep(0.2)
    SendToGame(keybinds[8])
    time.sleep(0.2)
    SendToGame(keybinds[8])
    time.sleep(0.2)
    SendToGame(keybinds[9])

def ReadKeyBinds():
    global keybinds
    with open('_keybinds.txt') as f:
        keybinds = [v for v in f.readline().split()]

    print("Key binds:", keybinds)
    if (len(keybinds) == 10):
        return True
    return False

def main():
    parser = argparse.ArgumentParser(description='BLARGH!')
    parser.add_argument('midifile', help='Path to a midi file')
    parser.add_argument('-t', dest='track', nargs='?', default=1, type=int, help='track number (default = 1)')
    #parser.add_argument('-ff', dest='play_all', nargs='?', default=0, type=int, help='play all tracks (if set to 1, ignores "track" parameter) (default = 0)')
    parser.add_argument('-s', dest='speed', nargs='?', default=1.0, type=float, help='speed modifier (default = 1.0)')
    parser.add_argument('-sync', dest='sync', nargs='?', default=0, type=int, help='play some notes before playing track (default = 0)')

    args = parser.parse_args()
    
    print("Opening", args.midifile)
    
    mid = MidiFile(args.midifile)
    trackNum = args.track
    #play_all = args.play_all
    speed = args.speed
    sync_count = args.sync    

    print("Total tracks: {}. Length: {}.".format(len(mid.tracks), mid.length))
    for i, trackx in enumerate(mid.tracks):
        print('Track {}: {}'.format(i, trackx.name.encode(sys.stdout.encoding, errors='replace')))

    if (trackNum < 0 or trackNum >= len(mid.tracks)):
        sys.exit("Error: Requested track not found.")
    if (speed < 0.001 or speed > 5):
        sys.exit("Error: speed < 0.001 or speed > 5.")
    if (sync_count < 0 or sync_count > 10):
        sys.exit("Error: sync_count < 0 or sync_count > 10.")

    k = ReadKeyBinds()
    if (not k):
        sys.exit("Error: something wrong with key bindings.")

    print("Pausing for a moment...")
    time.sleep(3)

    ResetInstrument()

    PlaySync(sync_count)

    #if (play_all):
    #    PlayAllFile(mid, speed)
    #else:
    track = mid.tracks[trackNum]
    PlayTrack(track, speed, mid.ticks_per_beat)        

    #for i, track in enumerate(mid.tracks):
    #    PlayTrack(track, mod)

if __name__ == "__main__":
    main()
