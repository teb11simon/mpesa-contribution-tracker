import json, os
from difflib import SequenceMatcher

class MemberRegistry:
    def __init__(self, filename="member_aliases.json"):
        self.filename = filename
        self.aliases = self._load() # Maps M-Pesa Name -> Excel Name

    def _load(self):
        if os.path.exists(self.filename):
            with open(self.filename, 'r') as f: return json.load(f)
        return {}

    def save_alias(self, mpesa_name, excel_name):
        self.aliases[mpesa_name.strip().upper()] = excel_name
        with open(self.filename, 'w') as f: json.dump(self.aliases, f)

    def get_verified_name(self, mpesa_name, member_list):
        m_name = mpesa_name.strip().upper()
        # 1. Check learned aliases first
        if m_name in self.aliases: return self.aliases[m_name]
        # 2. Try Fuzzy Match against your actual Excel member list
        best_match, highest_score = None, 0
        for member in member_list:
            score = SequenceMatcher(None, m_name, member['full_name'].upper()).ratio()
            if score > 0.8 and score > highest_score:
                highest_score, best_match = score, member['full_name']
        return best_match if best_match else mpesa_name