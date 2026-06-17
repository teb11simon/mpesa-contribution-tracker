import json, os

class SettingsManager:
    def __init__(self, filename="user_settings.json"):
        self.filename = filename
        self.settings = self._load()

    def _load(self):
        default = {"template_path": "", "output_dir": str(os.path.expanduser("~/Downloads"))}
        if os.path.exists(self.filename):
            with open(self.filename, 'r') as f: return {**default, **json.load(f)}
        return default

    def update(self, key, value):
        self.settings[key] = value
        with open(self.filename, 'w') as f: json.dump(self.settings, f)

    def get(self, key):
        return self.settings.get(key)