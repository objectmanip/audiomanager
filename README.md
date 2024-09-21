# About
Audiomanager automatically adjusts the volume of selected applications when both are playing audio based on custom configurations.
E.g. have music and a game running? Set spotify volume to 15% while the game stays at 100% or the other way round.

# Requirements
## On Windows:
### NirCmd
Requires NirCmd in the same directory. The files can be acquired here: https://www.nirsoft.net/utils/nircmd.html (downloads are at the end of the page). NirCmd is used for control volume levels.
### SoundVolumeView
Requires SoundVolumeView to change mute and listen settings for audiodevices for the hear-through option for headsets which don't support it. The files can be found here: https://www.nirsoft.net/utils/sound_volume_view.html (again, downloads are at the end of the page).

## On Linux
### pactl
Requires pactl to be installed for getting and setting audio application and device information and settings. Can usually be installed using your distros respective package manager, if it is not already installed.

# Features
## Control Output Volume on an Application Basis
Based on a set of simple .yaml files, applications will adjust their volume based on other applications which are currently playing audio (or are marked by the OS as playing audio). The lowest volume setting is chosen, when multiple referenced applications are running.

## Control Microphone Gain on an Application Basis (Currently Windows only)
Adjust the microphone gain based on applications running. E.g. Webex automatically adjust the windows microphone settings, this however does not revert back, and therefore TeamSpeak would usually overdrive. Setting the microphone gain, when TeamSpeak is started fixes the issue and that's what this option does.

## Allow Remote Mute and Play/Pause via API-Endpoints
Two api endpoints are available to toggle mute/unmute and play/pause.

## Settings for Headsets and Speakers
Currently, different setting can be made for up to two different audio outupt devices, e.g. a headset and speakers. In the *config.yaml*-file, the name of the **speaker** (as seen in your system audio settings) is stored. Based on that either speaker or headset settings are selected. For easily switching between audio devices, I recommend either SoundSwitch (on Windows) or a simple pactl-bash-script (on Linux).

# How to use


