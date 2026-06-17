class AnalyticsEngine:
    @staticmethod
    def summarize(income_list, expense_list):
        total_inc = sum(t['amount'] for t in income_list)
        total_exp = sum(t['amount'] for t in expense_list)
        
        # Group by category
        cat_totals = {}
        for t in income_list + expense_list:
            cat = t['category']
            cat_totals[cat] = cat_totals.get(cat, 0) + t['amount']
            
        return {
            "net": total_inc - total_exp,
            "income": total_inc,
            "expenses": total_exp,
            "breakdown": cat_totals
        }