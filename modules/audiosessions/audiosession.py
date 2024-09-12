class Process:
    def __init__(self, process_name, process_id) -> None:
        self.name = process_name
        self.id = process_id

class AudioSession:
    def __init__(self, name, process, process_id, state, current_volume):
        self.name = process
        self.Process = Process(process, process_id)
        self.State = state
        self.current_volume = current_volume
        self.DisplayName = process
