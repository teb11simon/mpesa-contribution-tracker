"""
Contribution Processor
The main controller that coordinates parsing, matching, and Excel generation.
"""

import logging
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path

from mpesa_parser import MpesaParser
from excel_generator import ExcelGenerator
from matching_engine import MatchingEngine

logger = logging.getLogger(__name__)

class ContributionProcessor:
    def __init__(self):
        self.parser = MpesaParser()
        self.generator = ExcelGenerator()
        self.matching_engine: Optional[MatchingEngine] = None

    def prepare_template(self, template_path: str):
        """Loads the template and prepares the generator/matching engine"""
        self.generator.load_template(template_path)
        members = self.generator._get_members_from_combined()
        self.matching_engine = MatchingEngine(members)
        return members

    def process_weekly_report(
        self,
        mpesa_pdf_path: Optional[str],
        pdf_password: Optional[str],
        cont_image_path: Optional[str],
        benev_image_path: Optional[str],
        miss_image_path: Optional[str],
        template_path: str,
        report_date: datetime,
        output_path: str,
        attendance: Optional[Dict] = None,
    ):
        """
        Processes everything for a single Sunday report.
        """
        logger.info(f"Starting processing for {report_date.strftime('%Y-%m-%d')}")
        
        # ... (Previous processing steps remain the same) ...
        # 1. Parse M-Pesa PDF
        transactions = []
        if mpesa_pdf_path:
            transactions = self.parser.parse_pdf(mpesa_pdf_path, password=pdf_password)
        
        # 2. Process Notes Images
        cash_entries = []
        combined_breakdown = {}
        if cont_image_path or benev_image_path or miss_image_path:
            from ocr_processor import OCRProcessor
            ocr = OCRProcessor()
            image_configs = [(cont_image_path, "Contribution"), (benev_image_path, "Benevolence"), (miss_image_path, "Missions")]
            for path, category in image_configs:
                if path:
                    entries_objs, breakdown = ocr.process_image(path)
                    for e in entries_objs:
                        e.category = category
                        cash_entries.append({'name': e.name, 'amount': e.amount, 'category': e.category})
                    for denom, total in breakdown.items():
                        combined_breakdown[denom] = combined_breakdown.get(denom, 0) + total
        cash_breakdown = combined_breakdown

        # 3. Load Template and Extract Members
        if not self.generator.workbook:
            self.generator.load_template(template_path)
        members = self.generator._get_members_from_combined()
        
        # 4. Initialize Matching Engine
        self.matching_engine = MatchingEngine(members)
        
        # 5. Map M-Pesa transactions to members (Paid In only)
        matched_mpesa = []
        unmatched_mpesa = []
        paid_out_transactions = []   # collected for Income & Exp tab

        for t in transactions:
            if t.transaction_type == "Paid Out":
                # Collect expenses for the Income & Exp tab
                paid_out_transactions.append({
                    'receipt_no':       t.receipt_no,
                    'date':             t.date,
                    'details':          t.details,
                    'amount':           t.amount,
                    'type':             'Paid Out',
                    'transaction_type': 'Paid Out',
                    'sender_name':      t.sender_name,
                    'category':         'Expense',
                })
                continue

            # --- Paid In handling (unchanged) ---
            match, score = self.matching_engine.find_match(t.sender_name, t.sender_phone, t.amount)
            category = "Contribution"
            details_lower = (t.details or "").lower()
            if "miss" in details_lower: category = "Missions"
            elif "benev" in details_lower: category = "Benevolence"
            t_dict = {'receipt_no': t.receipt_no, 'name': t.sender_name, 'phone': t.sender_phone, 'amount': t.amount, 'date': t.date, 'category': category, 'type': 'Paid In', 'transaction_type': 'Paid In', 'details': t.details}
            if match:
                t_dict['member_row'] = match['row_index']
                # Use the matched member's properly formatted first and last name
                first = str(match.get('first_name', '')).strip()
                last  = str(match.get('last_name',  '')).strip()
                t_dict['name'] = f"{first} {last}".title() if (first or last) else t.sender_name
                matched_mpesa.append(t_dict)
            else:
                unmatched_mpesa.append(t_dict)
        
        # 6. Map Cash Entries to members
        matched_cash = []
        unmatched_cash = []
        for c in cash_entries:
            match, score = self.matching_engine.find_match(c['name'], amount=c['amount'])
            if match:
                c['member_row'] = match['row_index']
                # Use properly formatted member name
                first = str(match.get('first_name', '')).strip()
                last  = str(match.get('last_name',  '')).strip()
                c['name'] = f"{first} {last}".title() if (first or last) else c['name']
                matched_cash.append(c)
            else: 
                unmatched_cash.append(c)

        # 7. Update member totals
        for m in members: m['amount'] = 0
        for t in matched_mpesa:
            for m in members:
                if m['row_index'] == t['member_row']: m['amount'] += t['amount']; break
        for c in matched_cash:
            for m in members:
                if m['row_index'] == c['member_row']: m['amount'] += c['amount']; break

        # 8. Generate Sunday Report
        self.generator.cash_breakdown = cash_breakdown
        self.generator.create_report_with_matches(report_date, output_path, matched_mpesa, unmatched_mpesa, matched_cash, unmatched_cash, members)
        
        # 9. Final Sync: Update Summary Dashboard, Combined Ledger, AND Attendance
        # Build the full flat list for Income & Exp tab:
        #   - all Paid In (matched + unmatched), each with receipt_no
        #   - all Paid Out expenses
        #   - cash entries (no receipt_no, so won't duplicate on re-run)
        all_transactions_for_ledger = (
            matched_mpesa +
            unmatched_mpesa +
            paid_out_transactions +
            [{'receipt_no': '', 'date': c.get('date'), 'details': c.get('name', ''),
              'amount': c.get('amount', 0), 'type': 'Paid In',
              'transaction_type': 'Paid In', 'category': c.get('category', 'Contribution')}
             for c in matched_cash + unmatched_cash]
        )

        self.generator.finalize_report(
            report_date, 
            matched_mpesa, 
            matched_cash, 
            members, 
            output_path,
            attendance=attendance,
            all_transactions=all_transactions_for_ledger
        )
        
        return {
            'matched_count': len(matched_mpesa) + len(matched_cash),
            'unmatched_count': len(unmatched_mpesa) + len(unmatched_cash),
            'output_path': output_path
        }
