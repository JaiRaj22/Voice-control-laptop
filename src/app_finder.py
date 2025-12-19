import os
import sys
import logging

class AppFinder:
    def __init__(self, logger=None):
        self.logger = logger or logging.getLogger(__name__)
        self.app_map = self._discover_apps()

    def _get_search_paths(self):
        """Returns a list of common application directories for Windows."""
        if sys.platform != "win32":
            return []

        # Standard environment variables for application locations
        user_profile = os.environ.get("USERPROFILE", "")
        app_data = os.environ.get("APPDATA", "")
        program_data = os.environ.get("PROGRAMDATA", "")
        public = os.environ.get("PUBLIC", "")

        # Key directories to scan
        paths = [
            os.path.join(user_profile, "Desktop"),
            os.path.join(public, "Desktop"),
            os.path.join(app_data, "Microsoft", "Windows", "Start Menu", "Programs"),
            os.path.join(program_data, "Microsoft", "Windows", "Start Menu", "Programs"),
        ]
        return [path for path in paths if path and os.path.isdir(path)]

    def _discover_apps(self):
        """Scans the search paths for .exe and .lnk files and builds the app map."""
        self.logger.info("Discovering installed applications...")
        app_map = {}
        search_paths = self._get_search_paths()

        if not search_paths:
            self.logger.warning("Could not find standard application directories. (Not on Windows?)")
            return app_map

        for path in search_paths:
            for root, _, files in os.walk(path):
                for file in files:
                    if file.lower().endswith((".exe", ".lnk")):
                        # Normalize the name for voice commands
                        app_name = os.path.splitext(file)[0].lower()
                        full_path = os.path.join(root, file)
                        
                        # Add to map if not already present (first found takes precedence)
                        if app_name not in app_map:
                            app_map[app_name] = full_path
        
        self.logger.info(f"Discovered {len(app_map)} applications.")
        return app_map

    def find_app(self, spoken_name):
        """
        Finds the path for an application based on a spoken name.
        Returns the full path or None if not found.
        """
        spoken_name = spoken_name.lower()
        
        # First, check for an exact match
        if spoken_name in self.app_map:
            return self.app_map[spoken_name]
            
        # If no exact match, check if any discovered app name contains the spoken name
        for app_name, path in self.app_map.items():
            if spoken_name in app_name:
                return path
        
        return None

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    finder = AppFinder()
    print("Discovered Applications:")
    for name, path in finder.app_map.items():
        print(f"  - {name}: {path}")
