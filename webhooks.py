from flask import Flask, request, Response
import time
from threading import Thread
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import pyautogui
import pythoncom

app = Flask(__name__)

@app.route('/playpause', methods=['POST'])
def playpause():
    pyautogui.press('playpause')
    return {}

@app.route('/togglemute', methods=['POST'])
def togglemute():
    pythoncom.CoInitialize()
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    if volume.GetMute():
        i = 0
    else:
        i = 1

    volume.SetMute(i, None)
    return {}

# app.run(host='0.0.0.0', port=5000, debug=False)
