import json, os, re, difflib
from typing import List, Dict, Optional, Tuple

class MatchingEngine:
    def __init__(self, members: List[Dict], aliases_file: Optional[str] = None):
        """
        Args:
            members: List of member dicts with 'first_name', 'last_name', 'phone', 'ministry'
            aliases_file: Path to JSON file for manual mappings. If None, defaults to 'member_aliases.json'
        """
        self.members = members
        self.aliases_file = aliases_file or "member_aliases.json"
        self.manual_aliases = self._load_aliases()
        self._normalize_member_list()

    def load_aliases(self):
        """Reload aliases from disk (useful after uploading a new aliases file)."""
        self.manual_aliases = self._load_aliases()

    def _load_aliases(self) -> Dict[str, str]:
        """Loads learned name mappings from JSON"""
        if os.path.exists(self.aliases_file):
            try:
                with open(self.aliases_file, 'r') as f:
                    data = json.load(f)
                    # Normalize keys to lowercase for consistent matching
                    return {k.lower(): v for k, v in data.items()}
            except Exception:
                return {}
        return {}

    def save_alias(self, query_name: str, member_full_name: str):
        """Persists a new manual mapping"""
        # Save as original Case in file, but we'll load as lower()
        self.manual_aliases[query_name.strip().lower()] = member_full_name.strip()
        try:
            # We want to keep the file structure simple
            with open(self.aliases_file, 'w') as f:
                json.dump(self.manual_aliases, f, indent=4)
        except Exception as e:
            print(f"Error saving alias: {e}")

    def _normalize_member_list(self):
        """Pre-calculate normalized names for faster matching"""
        for m in self.members:
            # Combine names and normalize
            first = str(m.get('first_name', '')).strip().lower()
            last = str(m.get('last_name', '')).strip().lower()
            m['full_name'] = f"{first} {last}".strip()
            m['reversed_name'] = f"{last} {first}".strip()

    def find_match(self, query_name: str, query_phone: Optional[str] = None, amount: Optional[float] = None) -> Tuple[Optional[Dict], float]:
        """
        Tries to find the best member match for a given name/phone.
        Returns (Matched Member Dict, Confidence Score 0.0-1.0)
        """
        if not query_name:
            return None, 0.0

        query_name = query_name.strip().lower()
        
        # 0. Check for explicit "Unmatch" directive
        if self.manual_aliases.get(query_name) == "--UNMATCHED--":
            return None, 0.0
        
        # 1. Manual Aliases & Special Cases (Learned from User)
        manual_match = self._check_manual_aliases(query_name, query_phone, amount)
        if manual_match:
            return manual_match, 1.0

        # 2. Phone Match (High Confidence)
        if query_phone:
            normalized_query_phone = self._normalize_phone(query_phone)
            for m in self.members:
                if m.get('phone'):
                    if self._normalize_phone(m['phone']) == normalized_query_phone:
                        return m, 1.0

        # 3. Exact Full Name Match
        for m in self.members:
            if query_name == m['full_name'] or query_name == m['reversed_name']:
                return m, 1.0

        # 3b. First Name Only Match (Fallback)
        # Often handwritten notes only use the first name
        first_name_matches = [
            m for m in self.members 
            if query_name == str(m.get('first_name', '')).strip().lower()
        ]
        if len(first_name_matches) == 1:
            return first_name_matches[0], 0.95

        # 4. Fuzzy Match using difflib
        best_match = None
        highest_score = 0.0
        
        for m in self.members:
            # Check against both normal and reversed name
            score = max(
                difflib.SequenceMatcher(None, query_name, m['full_name']).ratio(),
                difflib.SequenceMatcher(None, query_name, m['reversed_name']).ratio()
            )
            
            if score > highest_score:
                highest_score = score
                best_match = m
                
        # Confidence threshold
        if highest_score > 0.85:
            return best_match, highest_score
            
        return None, highest_score

    def _check_manual_aliases(self, name: str, phone: Optional[str], amount: Optional[float]) -> Optional[Dict]:
        """Handles specific learned mappings from the user"""
        name = name.strip().lower()
        
        # 1. Special hardcoded cases (logic-based)
        # Special NCBA Bank Case
        if "ncba bank" in name and amount == 2500:
            return self._find_by_full_name("arnold niragira")
            
        # Special Rita Wangare / Joy Muthoni Case
        if "rita wangare" in name:
            return self._find_by_full_name("joy muthoni")

        # 2. Dynamic Aliases (Loaded from JSON)
        if name in self.manual_aliases:
            return self._find_by_full_name(self.manual_aliases[name])

        # 3. Legacy Hardcoded Aliases (Transitioning to JSON)
        legacy_aliases = {
            "brian ajwala": "brian otieno",
            "frederick arogo": "fred arogo",
            "fredrick arogo": "fred arogo",
            "ayodeji aregbesola": "deji aregbesola",
            "aregbesolas": "deji aregbesola",
            "opwondis": "james opwondi",
            "suzie": "susan njoki njoro",
            "myra kadenge": "myra minayo kadenge",
            "jackline mutuku": "jackline mwende mutuku",
            "jenipher angira": "jenipher angira",
            "john gathenya": "john maingi",
            "maurice mbogo": "maurice odongo mbogo"
        }
        
        if name in legacy_aliases:
            return self._find_by_full_name(legacy_aliases[name])
            
        return None

    def _find_by_full_name(self, full_name: str) -> Optional[Dict]:
        """Helper to find a member by their full normalized name"""
        full_name = full_name.strip().lower()
        for m in self.members:
            if m['full_name'] == full_name or m['reversed_name'] == full_name:
                return m
        return None

    def _normalize_phone(self, phone: str) -> str:
        """Removes non-digits and handles 254 prefix"""
        digits = re.sub(r'\D', '', str(phone))
        if len(digits) == 9 and digits.startswith('7'):
            digits = '254' + digits
        elif len(digits) == 10 and digits.startswith('0'):
            digits = '254' + digits[1:]
        return digits

def get_engine(members: List[Dict]) -> MatchingEngine:
    return MatchingEngine(members)
