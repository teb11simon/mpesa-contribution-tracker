"""
M-Pesa Statement Parser
Extracts transaction data from M-Pesa statements (PDF/CSV)
"""

import re
import csv
import logging
from datetime import datetime
from typing import List, Dict, Optional
from dataclasses import dataclass

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class Transaction:
    """Represents a single M-Pesa transaction"""
    receipt_no: str
    date: datetime
    details: str
    amount: float
    transaction_type: str  # 'Paid In' or 'Paid Out'
    sender_name: str = ""
    sender_phone: str = ""
    balance: Optional[float] = None

class MpesaParser:
    """Parser for M-Pesa statements (Safaricom format)"""

    def __init__(self):
        self.transactions: List[Transaction] = []

    def parse_pdf(self, pdf_path: str, password: Optional[str] = None) -> List[Transaction]:
        """
        Parse M-Pesa statement from PDF file.
        Supports encrypted PDFs using the provided password (usually National ID).
        """
        try:
            import pdfplumber
        except ImportError:
            raise ImportError("pdfplumber required for PDF parsing. Install with: pip install pdfplumber")

        transactions = []
        
        try:
            with pdfplumber.open(pdf_path, password=password) as pdf:
                for page in pdf.pages:
                    # Extract table using Safaricom-specific settings
                    # Safaricom tables usually have clear vertical lines
                    table = page.extract_table({
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_tolerance": 3,
                    })
                    
                    if not table:
                        # Fallback to text parsing if table extraction fails
                        text = page.extract_text()
                        if text:
                            transactions.extend(self._parse_raw_text(text))
                        continue

                    # Process table rows
                    # Header expected: Receipt No. | Completion Time | Details | Transaction Status | Paid In | Paid Out | Balance
                    for row in table:
                        if not row or "Receipt No" in str(row[0]):
                            continue
                        
                        trans = self._parse_table_row(row)
                        if trans:
                            transactions.append(trans)

        except Exception as e:
            logger.error(f"Error parsing PDF {pdf_path}: {e}")
            raise

        self.transactions = transactions
        return transactions

    def _parse_table_row(self, row: List[Optional[str]]) -> Optional[Transaction]:
        """
        Parses a single row from the Safaricom PDF table.
        Standard Safaricom PDF columns:
        0: Receipt No.
        1: Completion Time
        2: Details
        3: Transaction Status
        4: Paid In
        5: Paid Out
        6: Balance
        """
        try:
            if len(row) < 7:
                return None

            receipt_no = (row[0] or "").strip()
            time_str = (row[1] or "").strip()
            details = (row[2] or "").strip()
            paid_in = (row[4] or "").strip()
            paid_out = (row[5] or "").strip()
            balance_str = (row[6] or "").strip()

            if not receipt_no or not time_str:
                return None

            # Parse amount and type
            amount = 0.0
            trans_type = ""
            if paid_in and paid_in != "-":
                amount = float(paid_in.replace(",", ""))
                trans_type = "Paid In"
            elif paid_out and paid_out != "-":
                amount = float(paid_out.replace(",", ""))
                trans_type = "Paid Out"
            else:
                return None

            # Parse date
            # Safaricom format: 2026-05-03 14:30:15
            try:
                date = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                # Try fallback format if needed
                date = datetime.now()

            # Extract name and phone from details
            name, phone = self._extract_info_from_details(details)

            return Transaction(
                receipt_no=receipt_no,
                date=date,
                details=details,
                amount=amount,
                transaction_type=trans_type,
                sender_name=name,
                sender_phone=phone,
                balance=float(balance_str.replace(",", "")) if balance_str else None
            )

        except (ValueError, IndexError) as e:
            logger.debug(f"Row parsing failed: {e} | Row: {row}")
            return None

    def _extract_info_from_details(self, details: str) -> (str, str):
        """
        Extracts sender name and phone from the details string.
        Examples:
        - "Customer Transfer of Funds from 254712345678 - JOHN DOE"
        - "Received from 254712345678 - JOHN DOE"
        - "Pay Bill to 123456 - CHURCH ACCOUNT"
        """
        name = ""
        phone = ""

        # Pattern for "from/to [NUMBER] - [NAME]"
        # Matches 254... or any digit sequence
        match = re.search(r'(?:from|to)\s+(\d+)\s+-\s+(.+)', details, re.IGNORECASE)
        if match:
            phone = match.group(1).strip()
            name = match.group(2).strip()
        else:
            # Fallback patterns if format differs
            # Look for 10-12 digit phone number
            phone_match = re.search(r'(254\d{9}|\d{10})', details)
            if phone_match:
                phone = phone_match.group(1)
            
            # Name usually follows a dash or comes at the end
            name_match = re.search(r'-\s+([A-Z\s]+)$', details)
            if name_match:
                name = name_match.group(1).strip()

        # Final Fallback: If no name found by regex, use the details string itself
        if not name and details:
            # Clean up the details string (remove common prefixes)
            name = re.sub(r'^(?:Funds received from|Customer Transfer of Funds from|Received from|Pay Bill to)\s*(?:-)?\s*', '', details, flags=re.IGNORECASE)
            # Remove phone masks like "07******436" or "2547******436" at the beginning
            name = re.sub(r'^[0-9\*]+\s*', '', name)
            # Remove any remaining phone numbers or common noise
            name = re.sub(r'254\d{9}|\d{10,12}', '', name).strip()
            # Clean up trailing/leading dashes or special chars
            name = re.sub(r'^[-:\s]+|[-:\s]+$', '', name)
            
            # If it's still empty or just noise, use full details
            if not name or len(name) < 2:
                name = details.strip()

        return name, phone

    def _parse_raw_text(self, text: str) -> List[Transaction]:
        """
        Fallback parser using regex on raw text if table extraction fails.
        """
        transactions = []
        # Regex for Safaricom line: Receipt | Date Time | Details | Status | Paid In | Paid Out | Balance
        # This is harder to get right across pages, hence why table extraction is preferred.
        pattern = r'([A-Z0-0]{10})\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+(.+?)\s+Completed\s+([\d,.-]+)\s+([\d,.-]+)\s+([\d,.]+)'
        
        for match in re.finditer(pattern, text):
            receipt = match.group(1)
            time_str = match.group(2)
            details = match.group(3)
            paid_in = match.group(4)
            paid_out = match.group(5)
            balance = match.group(6)
            
            try:
                date = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                amount = 0.0
                trans_type = ""
                
                if paid_in != "-":
                    amount = float(paid_in.replace(",", ""))
                    trans_type = "Paid In"
                else:
                    amount = float(paid_out.replace(",", ""))
                    trans_type = "Paid Out"
                
                name, phone = self._extract_info_from_details(details)
                
                transactions.append(Transaction(
                    receipt_no=receipt,
                    date=date,
                    details=details,
                    amount=amount,
                    transaction_type=trans_type,
                    sender_name=name,
                    sender_phone=phone,
                    balance=float(balance.replace(",", ""))
                ))
            except ValueError:
                continue
                
        return transactions

    def parse_csv(self, csv_path: str) -> List[Transaction]:
        """Parse M-Pesa statement from CSV file"""
        transactions = []

        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                transaction = self._parse_csv_row(row)
                if transaction:
                    transactions.append(transaction)

        self.transactions = transactions
        return transactions

    def _parse_csv_row(self, row: Dict) -> Optional[Transaction]:
        """Parse a single CSV row into a Transaction"""
        try:
            # Try common column names
            receipt = row.get('Receipt No.') or row.get('Receipt') or row.get('id')
            date_str = row.get('Completion Time') or row.get('Date') or row.get('date')
            details = row.get('Details') or row.get('Description') or row.get('details')
            paid_in = row.get('Paid In') or row.get('Amount') or row.get('amount')
            paid_out = row.get('Paid Out')
            balance_str = row.get('Balance') or row.get('balance')

            if not date_str or (not paid_in and not paid_out):
                return None

            # Parse amount and type
            amount = 0.0
            trans_type = "Received"
            if paid_in and str(paid_in) != "-":
                amount = float(str(paid_in).replace(",", ""))
                trans_type = "Paid In"
            elif paid_out and str(paid_out) != "-":
                amount = float(str(paid_out).replace(",", ""))
                trans_type = "Paid Out"

            # Parse date
            if isinstance(date_str, str):
                try:
                    date = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    try:
                        date = datetime.strptime(date_str, "%d/%m/%Y %H:%M")
                    except ValueError:
                        date = datetime.now()
            else:
                date = date_str

            # Extract info from details if available
            name, phone = self._extract_info_from_details(details or "")

            return Transaction(
                receipt_no=receipt or "N/A",
                date=date,
                details=details or "",
                amount=amount,
                transaction_type=trans_type,
                sender_name=name,
                sender_phone=phone,
                balance=float(str(balance_str).replace(",", "")) if balance_str else None
            )

        except (ValueError, KeyError, AttributeError):
            return None

    def to_dict_list(self) -> List[Dict]:
        """Convert transactions to list of dictionaries for easier JSON/DataFrame use"""
        return [
            {
                'receipt_no': t.receipt_no,
                'date': t.date.strftime('%Y-%m-%d %H:%M:%S'),
                'details': t.details,
                'amount': t.amount,
                'type': t.transaction_type,
                'sender_name': t.sender_name,
                'sender_phone': t.sender_phone,
                'balance': t.balance
            }
            for t in self.transactions
        ]
