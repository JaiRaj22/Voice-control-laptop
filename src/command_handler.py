import os
import subprocess
import sys
import logging
from app_finder import AppFinder

class CommandHandler:
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        self.app_finder = AppFinder(logger=self.logger)
        self.commands = {
            "open": "open_application",
            "volume up": "volume_up",
            "volume down": "volume_down",
            "mute": "mute_volume",
            "shutdown": "shutdown",
            "restart": "restart",
            "sleep": "sleep",
        }
        self._volume_interface = self._get_volume_interface()

    def _get_volume_interface(self):
        if sys.platform == "win32":
            try:
                from pycaw.pycaw import AudioUtilities
                speakers = AudioUtilities.GetSpeakers()
                return speakers.EndpointVolume
            except Exception as e:
                self.logger.error(f"Failed to initialize volume control: {e}")
                return None
        return None

    def execute_command(self, command_text):
        if not command_text:
            return

        command_text = command_text.lower()
        for keyword, action in self.commands.items():
            if command_text.startswith(keyword):
                if isinstance(action, str):
                    method = getattr(self, action)
                else:
                    method = action

                if keyword == "open":
                    app_name = command_text[len(keyword):].strip()
                    method(app_name)
                else:
                    method()
                return

    def open_application(self, app_name):
        app_path = self.app_finder.find_app(app_name)
        if app_path:
            try:
                os.startfile(app_path)
                self.logger.info(f"Opening {app_name} from {app_path}")
            except Exception as e:
                self.logger.error(f"Could not open '{app_name}': {e}")
        else:
            self.logger.warning(f"Application '{app_name}' not found.")

    def volume_up(self):
        if not self._volume_interface:
            self.logger.warning("Volume control is not supported on this OS.")
            return
        current_volume = self._volume_interface.GetMasterVolumeLevelScalar()
        new_volume = min(1.0, current_volume + 0.1)
        self._volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
        self.logger.info(f"Volume set to {new_volume * 100:.0f}%")

    def volume_down(self):
        if not self._volume_interface:
            self.logger.warning("Volume control is not supported on this OS.")
            return
        current_volume = self._volume_interface.GetMasterVolumeLevelScalar()
        new_volume = max(0.0, current_volume - 0.1)
        self._volume_interface.SetMasterVolumeLevelScalar(new_volume, None)
        self.logger.info(f"Volume set to {new_volume * 100:.0f}%")

    def mute_volume(self):
        if not self._volume_interface:
            self.logger.warning("Volume control is not supported on this OS.")
            return
        self._volume_interface.SetMute(not self._volume_interface.GetMute(), None)
        self.logger.info("Mute toggled")

    def shutdown(self, confirmation_callback=None):
        if confirmation_callback and not confirmation_callback():
            self.logger.info("Shutdown cancelled.")
            return

        self.logger.info("Shutting down...")
        os.system("shutdown /s /t 1")

    def restart(self, confirmation_callback=None):
        if confirmation_callback and not confirmation_callback():
            self.logger.info("Restart cancelled.")
            return

        self.logger.info("Restarting...")
        os.system("shutdown /r /t 1")

    def sleep(self):
        self.logger.info("Putting the computer to sleep...")
        os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
