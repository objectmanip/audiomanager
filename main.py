"""
    Sets volume of programs according to the settings
"""
import os
import sys
import time
import yaml
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from threading import Thread
from PyQt5.QtGui import *
from PyQt5.QtWidgets import QApplication, QAction, QMenu, QSystemTrayIcon
from PyQt5.QtCore import QFile, QTextStream
import subprocess
import sounddevice
import logging
from logging.handlers import RotatingFileHandler
from waitress import serve
import requests
from webhooks import app as webhook_app
import win32api
import win32gui


log = logging.getLogger("audiomanager")
log.setLevel(logging.INFO)  

file_handler = RotatingFileHandler("./logs/audiomanager.log", mode="w", maxBytes= 2*1024*1024, backupCount=3)
file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
log.addHandler(file_handler)
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))
log.addHandler(console_handler)
CURRENT_DIRECTORY = os.path.dirname(os.path.realpath(__file__))
os.makedirs('./config', exist_ok=True)
os.makedirs('./logs', exist_ok=True)
os.makedirs('./profiles', exist_ok=True)

class AudioManager:
    def __init__(self):
        # set hear through to off for consistent base settings
        self.config_path = './config/config.yaml'
        self.log_path = './logs'
        self.profiles_path = './profiles'
        self.config = {}
        self.hear_through_enabled = False
        self.__load_config()
        subprocess.call(f'SoundVolumeView.exe /SetListenToThisDevice "{self.config["microphone_name"]}" 0', shell=True)
        self.volume_threads = {}
        self.keep_alive = True
        log.info("Initalized")
        self.__get_audio_sessions()
        Thread(target=self.__auto_volume, daemon=False).start()
        Thread(target=self.webhooker, daemon=True).start()
        self.tray_menu()

    def webhooker(self):
        # subprocess.call("waitress-serve --listen *:5000 wsgi:app")
        serve(webhook_app, listen=f"*:{self.config['port']}")

    def __toggle_hear_through(self):
        try:
            if not self.hear_through_enabled: # and self.__get_current_audio_device() == 'headset':
                log.info(f"Set hear through: on")
                subprocess.call(f'SoundVolumeView.exe /SetListenToThisDevice "TonorMikrofon" 1', shell=True)
            else:
                log.info(f"Set hear through: off")
                subprocess.call(f'SoundVolumeView.exe /SetListenToThisDevice "TonorMikrofon" 0', shell=True)
        except:
            log.info("Could not toggle hear through, wrong ground state.")
        else:
            self.hear_through_enabled = not self.hear_through_enabled

    def __auto_volume(self):
        """
        manages the combination of volume profiles and sets the volume
        :return:
        """
        import pythoncom
        pythoncom.CoInitialize()

        while self.keep_alive:
            self.__load_config()
            self.__get_audio_sessions()
            # self.__set_capture_card_volume()
            current_audio_device = self.__get_current_audio_device()

            mic_gain = self.config['microphone_gain']['base']
            set_microphone_application = False
            for microphone_application in self.mic_profiles:
                if self.__match_processes(microphone_application):
                    if self.dev_log: log.info(f"Mic-application found: {microphone_application}")
                    if self.mic_profiles[microphone_application] is None:
                        set_microphone_application = True
                        continue
                    if self.mic_profiles[microphone_application] < mic_gain:
                        mic_gain = self.mic_profiles[microphone_application]
                        set_microphone_application = True

            if self.hear_through_enabled:
                microphone_offset = self.config['microphone_gain']['hear_through_offset']
            else:
                microphone_offset = 0

            if mic_gain is not None:
                self.__set_microphone_gain(mic_gain-microphone_offset)
            self.__set_capture_card_volume()
            for application_to_set in self.volume_profiles:
                # skip if application is not running
                target_session = self.__match_processes(application_to_set)

                if not target_session:
                    continue
                # base volume to set to if no app is found, will be set to the lowest matched value
                target_volume = self.volume_profiles[application_to_set]["standard"][current_audio_device]
                profile_applications = self.volume_profiles[application_to_set]
                for application_to_watch in profile_applications:
                    # skip if application is not running
                    session = self.__match_processes(application_to_watch)
                    # check if hear through is activated
                    if application_to_watch == "hear_through" and self.hear_through_enabled:
                        pass
                    elif not session and not self.config['reset_volume_sessions']:
                        continue
                    elif self.config['reset_volume_sessions']:
                        target_volume = 1
                        break
                    elif not self.is_profile_active(application_to_watch):
                        continue
                    elif (not session.State or session.SimpleAudioVolume.GetMute()) and self.config['check_watched_application_state']:
                        # if session is not playing audio or muted
                        continue

                    # get the lowest volume for application
                    if profile_applications[application_to_watch][current_audio_device] < target_volume or \
                            profile_applications["standard"][current_audio_device] == target_volume == 0:
                        target_volume = profile_applications[application_to_watch][current_audio_device]
                    else:
                        log.debug(f"Skipping {application_to_set}, {application_to_watch}")

                # queue thread for setting volume
                application_thread = Thread(target=self.__set_app_volume,
                                            args=(target_session, target_volume, session),
                                            daemon=True)
                self.__queue__set_app_volume(application_thread, application_to_set)

            time.sleep(1)

    def is_profile_active(self, application_profile):
        for key in self.profile_applications.keys():
            # if application profile (e.g. 'communication' for 'discord', 'atmgr', etc) is not active:
            # check if application is in said profile
            if application_profile in self.profile_applications[key] and not self.config['profiles'][key]:
                return False

        return True

    def __get_audio_sessions(self):
        """
        sets self.audio_sessions to a list of lists with (process_name, process_id) as elements
        :return:
        """
        self.audio_sessions = AudioUtilities.GetAllSessions()
        if self.config['list_active_audio_sessions']:
            log.info("Active Audio Sessions")
            for session in self.audio_sessions:
                try:
                    log.info(session.Process.name())
                except AttributeError:
                    pass

    def __get_current_audio_device(self):
        """

        :return:
        """
        sounddevice._terminate()
        sounddevice._initialize()

        if self.config['speakername'] in str(sounddevice.query_devices(sounddevice.default.device[1])):
            try:
                requests.post(self.config['urls']['homeassistant']['toggle_off'])
            except:
                return 'headset'
            else:
                return 'speaker'
        else:
            try:
                requests.post(self.config['urls']['homeassistant']['toggle_on'])
            except:
                return 'headset'
            else:
                return 'headset'

    def __get_current_window(self):
        pass

    def __get_current_volume(self, audio_session) -> float:
        """
        returns the volume of the application
        :param audio_session:
        :return:
        """
        pass

    def __load_config(self):
        """
        :return:
        """
        while True:
            try:
                with open(self.config_path, "r", encoding="utf-8") as cf:
                    self.config = yaml.load(cf, yaml.Loader)
                
                with open(os.path.join(self.profiles_path, "profiles.yaml"), "r", encoding="utf-8") as pf:
                    self.volume_profiles = yaml.load(pf, yaml.Loader)

                self.profile_applications = {}
                for file in [os.path.join(self.profiles_path, file) 
                             for file in os.listdir(self.profiles_path) 
                             if file.startswith('profiles_') and not "microphone" in file]:
                    matched_profiles = []
                    with open(file, "r", encoding="utf-8") as gpf:
                        profiles = yaml.load(gpf, yaml.Loader)

                    profile_id = os.path.basename(file).replace("profiles_", '').replace(".yaml",'')
                    if profile_id not in self.config['profiles']:
                        self.config['profiles'][profile_id] = True

                    self.profile_applications[profile_id] = profiles.keys()
                    for profile in profiles:
                        # add profiles for each subprofile
                        if profile_id in self.volume_profiles:
                            self.volume_profiles[profile] = self.volume_profiles[profile_id].copy()
                            self.volume_profiles[profile]["standard"] = {"headset": profiles[profile],
                                                                         "speaker": profiles[profile]}
                        # add suprofiles for each profile if the profile id is present in a volume setting
                        # iterate through profiles which are used
                        for volpro in self.volume_profiles.items():
                            if profile_id in volpro[1].keys():
                                setting = volpro[1][profile_id]

                                self.volume_profiles[volpro[0]][profile] = setting
                                matched_profiles.append([volpro[0], profile_id])

                    # print(self.volume_profiles)
                    if profile_id in self.volume_profiles:
                        self.volume_profiles.pop(profile_id)
                    for mp in matched_profiles:
                        try:
                            del(self.volume_profiles[mp[0]][mp[1]])
                        except KeyError:
                            pass
                # print(self.volume_profiles)

                with open(os.path.join(self.profiles_path, "profiles_microphone.yaml"), "r", encoding="utf-8") as pf:
                    self.mic_profiles = yaml.load(pf, yaml.Loader)
                if 'microphone' not in self.config['profiles']:
                    self.config['profiles']['microphone'] = True

                self.dev_log = self.config['dev_log']

            except AttributeError as err:
                log.info("Failed to load config.yaml.", err)
                time.sleep(5)
            else:
                break
            self.__save_config()

    def __match_processes(self, process: str):
        """

        :return:
        """
        named_sessions = []

        for session in self.audio_sessions:
            try:
                try:
                    session_name = session.Process.name()
                except:
                    session_name = session.DisplayName
                named_sessions.append([session, session_name])
            except AttributeError:
                pass

        matched_sessions = [session[0] for session in named_sessions if
                            process.lower() in session[1].lower()]

        if len(matched_sessions) > 1:
            for session in matched_sessions:
                if session.State == 1:
                    # if multiple sessions, return the active session or the last session
                    return session
            else:
                return session
        elif len(matched_sessions) == 0:
            return False
        else:
            # if only one session, return that session
            return matched_sessions[0]

    def __open_settings(self, file: str):
        log.info(f"Opening {file}")
        subprocess.call(f"notepad.exe {os.path.join(self.profiles_path, file)}", shell=True)

    def __save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as cf:
            yaml.dump(self.config, cf)

    def __queue__set_app_volume(self, thread_element, thread_name):
        """

        :param thread_element:
        :param thread_name:
        :return:
        """
        for running_thread in self.volume_threads.copy():
            if not self.volume_threads[running_thread].is_alive():
                self.volume_threads[running_thread].join()
                self.volume_threads.pop(running_thread)

        if thread_name in self.volume_threads:
            return
        else:
            self.volume_threads[thread_name] = thread_element
            self.volume_threads[thread_name].start()

    def __quit(self):
        self.keep_alive = False
        self.app.quit()
        exit()

    def __restart(self):
        os.execv(sys.executable, ['python'] + sys.argv)

    def __set_app_volume(self, audio_session, target_volume, watched_session = None):
        """

        :param audio_session:
        :param target_volume:
        :param watched_session:
        :return:
        """
        # get audiosessions
        # volume_element = audio_session._clt.QueryInterface(ISimpleAudioVolume)
        try:
            current_volume = round(audio_session.SimpleAudioVolume.GetMasterVolume(), 3)
        except:
            return
        if current_volume == target_volume or not self.config['active']:
            return


        volume_steps = (target_volume - current_volume)/self.config['transition_length']
        if volume_steps < 0:
            volume_steps *= (self.config['transition_length']/2)

        try:
            session_name = audio_session.Process.name()
        except:
            session_name = audio_session.DisplayName
        # if watched_session:
        #     log.info(f"Setting volume for {session_name} to {target_volume*100}% ({watched_session.Process.name()}, Active: {watched_session.State}, Muted: {watched_session.SimpleAudioVolume.GetMute()}).")
        # else:
        log.info(f"Setting volume for {session_name} to {target_volume*100}%.")

        while (current_volume + volume_steps < target_volume and current_volume < target_volume) or \
                (current_volume + volume_steps > target_volume and current_volume > target_volume):
            audio_session.SimpleAudioVolume.SetMasterVolume(current_volume + volume_steps, None)
            current_volume += volume_steps
            time.sleep(.05)
        audio_session.SimpleAudioVolume.SetMasterVolume(target_volume, None)

    def __set_capture_card_volume(self):
        self.__get_audio_sessions()
        process = self.__match_processes(process="svchost.exe")
        volume = self.config['capture_card']['mode_on'] if self.config['capture_card']['state'] else self.config['capture_card']['mode_off']
        self.__set_app_volume(process, int(volume))

    def __set_microphone_gain(self, gain):
        if self.dev_log: log.info(f"Setting mic gain to {gain}")
        # os.system(f"nircmdc.exe loop 1 250 setsysvolume {gain} default_record")
        try:
            subprocess.call(f"nircmdc.exe loop 1 250 setsysvolume {gain} default_record", shell=True)
        except PermissionError:
            log.info("Error: Microphone access denied")

    def __toggle_settings(self, para: str):
        try:
            self.config[para] = not self.config[para]
            log.info(f"Set {para} to: {self.config[para]}")
        except:
            self.config['profiles'][para] = not self.config['profiles'][para]
            log.info(f"Set {para} to: {self.config['profiles'][para]}")

        self.__save_config()

    def tray_menu(self):
        log.info("Launching Tray Icon")
        self.app = QApplication([])
        self.icon = QIcon("icons/icon.ico")
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(self.icon)
        self.tray.setVisible(True)
        self.app.setQuitOnLastWindowClosed(False)

        # Tray-Menu
        menu = QMenu()

        menu.addAction('Toggle audiomanager', lambda: self.__toggle_settings('active'))
        menu.addAction('Toggle hear-through', lambda: self.__toggle_hear_through())
        menu.addAction('Toggle capture card audio', lambda: self.__toggle_settings('capture_card'))
        menu.addAction('Toggle application check', lambda: self.__toggle_settings('check_watched_application_state'))
        menu.addSeparator()
        for file in [file for file in os.listdir(self.profiles_path) if file.startswith('profile') and file.endswith('.yaml')]:
            audio_option = file.replace('profiles_','').replace('.yaml','') if file != 'profiles.yaml' else 'profiles'
            action = QAction(f'Toggle {audio_option}', menu)
            action.triggered.connect(lambda checked, arg=audio_option: self.__toggle_settings(arg))
            menu.addAction(action)
        menu.addSeparator()
        menu.addAction('Reset', lambda: self.__toggle_settings('reset_volume_sessions'))
        menu.addSeparator()
        for file in [file for file in os.listdir(self.profiles_path) if file.startswith('profile') and file.endswith('.yaml')]:
            audio_option = file.replace('profiles_','').replace('.yaml','') if file != 'profiles.yaml' else 'profiles'
            action = QAction(f'Edit {audio_option}', menu)
            action.triggered.connect(lambda checked, arg=file: self.__open_settings(arg))
            menu.addAction(action)

        menu.addSeparator()

        option_close = QAction("Close")
        option_close.triggered.connect(self.__quit)
        menu.addAction(option_close)

        self.tray.setContextMenu(menu)
        log.info("Tray started.")
        self.app.exec_()

def main():
    app = AudioManager()

if __name__ == "__main__":
    if "--audiosessions" in sys.argv:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            try: print(session.Process.name(), session.State)
            except: print(session.DisplayName, session.State)
    else:
        main()
