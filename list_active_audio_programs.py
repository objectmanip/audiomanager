from pycaw.pycaw import AudioUtilities

def _get_audio_sessions():
    """
    sets self.audio_sessions to a list of lists with (process_name, process_id) as elements
    :return:
    """
    audio_sessions = AudioUtilities.GetAllSessions()
    print("Active Audio Sessions")
    for session in audio_sessions:
        try:
            print(session.Process.name())
        except AttributeError:
            pass

if __name__ == '__main__':
    _get_audio_sessions()