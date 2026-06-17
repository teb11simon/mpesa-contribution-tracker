import json
import os

# Hardcoded rules for instant automation
AUTO_RULES = {
    "Transaction Charge": "Transaction Charge",
    "M-PESA Withdrawal Charge": "Transaction Charge",
    "M-PESA Transfer Charge": "Transaction Charge",
}

class CategoryMemory:
    def __init__(self, filename="app_memory.json"):
        self.filename = filename
        self.data = self._load()

    def _load(self):
        if os.path.exists(self.filename):
            try:
                with open(self.filename, 'r') as f:
                    return json.load(f)
            except: pass
        return {"days": {}, "senders": {}}

    def save(self):
        with open(self.filename, 'w') as f:
            json.dump(self.data, f)

    def learn(self, transaction, category):
        if category == "__IGNORE__": return
        
        # Learn by Day (0=Mon, 6=Sun)
        day_key = str(transaction.date.weekday())
        day_stats = self.data["days"].get(day_key, {})
        day_stats[category] = day_stats.get(category, 0) + 1
        self.data["days"][day_key] = day_stats

        # Learn by Sender
        if transaction.sender_name:
            sender_key = transaction.sender_name.strip().upper()
            sender_stats = self.data["senders"].get(sender_key, {})
            sender_stats[category] = sender_stats.get(category, 0) + 1
            self.data["senders"][sender_key] = sender_stats
        self.save()

    def suggest(self, transaction):
        # 1. Rules first (Automated)
        for keyword, category in AUTO_RULES.items():
            if keyword.lower() in transaction.details.lower():
                return category, True # True means it's an auto-match

        # 2. Sender Memory
        if transaction.sender_name:
            sender_key = transaction.sender_name.strip().upper()
            stats = self.data["senders"].get(sender_key)
            if stats:
                return max(stats, key=stats.get), False

        # 3. Day Memory (Wednesday/Sunday logic)
        day_key = str(transaction.date.weekday())
        stats = self.data["days"].get(day_key)
        if stats:
            return max(stats, key=stats.get), False
        
        return "Uncategorized", False