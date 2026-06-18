import json, os
from pathlib import Path

class SettingsManager:
    def __init__(self, church_slug: str = "nairobi-icc", filename: str = "user_settings.json"):
        """
        Per-church settings manager.
        
        Args:
            church_slug: The unique slug for the church (e.g., 'nairobi-icc')
            filename: The settings filename (defaults to user_settings.json)
        """
        self.church_slug = church_slug
        self.filename = filename
        
        # Ensure church directory exists
        self.church_dir = Path("churches") / church_slug
        self.church_dir.mkdir(parents=True, exist_ok=True)
        
        # Full path to settings file
        self.settings_path = self.church_dir / filename
        self.settings = self._load()

    def _load(self):
        default = {
            "template_path": "",
            "output_dir": str(os.path.expanduser("~/Downloads"))
        }
        if self.settings_path.exists():
            with open(self.settings_path, 'r') as f:
                return {**default, **json.load(f)}
        return default

    def update(self, key, value):
        self.settings[key] = value
        with open(self.settings_path, 'w') as f:
            json.dump(self.settings, f)

    def get(self, key):
        return self.settings.get(key)
    
    def get_church_dir(self) -> Path:
        """Returns the church-specific directory path."""
        return self.church_dir
    
    def get_aliases_path(self) -> Path:
        """Returns the path to the church-specific member_aliases.json."""
        return self.church_dir / "member_aliases.json"