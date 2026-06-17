"""
OCR Module for Processing Handwritten Contribution Notes
Extracts names and amounts from images using OCR
"""

import re
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass
from pathlib import Path


@dataclass
class ContributionEntry:
    """Represents a contribution entry from handwritten notes"""
    name: str
    amount: float
    category: str  # 'Contribution', 'Benevolence', 'Missions'
    ministry: Optional[str] = None
    notes: str = ""


class OCRProcessor:
    """Process images to extract contribution data"""

    def __init__(self, use_google_vision: bool = False):
        """
        Initialize OCR processor

        Args:
            use_google_vision: If True, use Google Cloud Vision API (requires API key)
                              If False, use Tesseract (local)
        """
        self.use_google_vision = use_google_vision
        self.entries: List[ContributionEntry] = []

    def process_image(self, image_path: str) -> Tuple[List[ContributionEntry], Dict[float, float]]:
        """
        Process an image to extract contribution entries and cash breakdown
        """
        if self.use_google_vision:
            text = self._extract_text_google_vision(image_path)
        else:
            text = self._extract_text_tesseract(image_path)

        entries = self._parse_contribution_text(text)
        breakdown = self._parse_denominations(text)
        self.entries = entries
        self.breakdown = breakdown
        return entries, breakdown

    def _parse_denominations(self, text: str) -> Dict[float, float]:
        """
        Extracts cash denominations from text.
        Handles formats: "1000 x 5", "1000x5", "1000 x5", "200X3", "100 % 1" (OCR misread)
        """
        breakdown = {}
        # Allow no spaces around separator, and accept % as OCR misread of x
        pattern = r'(\d[\d,]*)\s*(?:[xX\*%+\-\/\u00D7]\s*|\s+)(\d+)\s*(?:[=:]\s*([\d,]+))?'

        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
            match = re.search(pattern, line)
            if match:
                try:
                    denom = float(match.group(1).replace(',', ''))
                    count = int(match.group(2))
                    # Sanity check — denominations should be round numbers ≥ 1
                    # and counts should be reasonable (< 1000)
                    if denom < 1 or count < 1 or count > 999:
                        continue
                    total = float(match.group(3).replace(',', '')) if match.group(3) else denom * count
                    breakdown[denom] = breakdown.get(denom, 0) + total
                except (ValueError, TypeError):
                    continue
        return breakdown

    def _extract_text_tesseract(self, image_path: str) -> str:
        """Extract text using Tesseract OCR"""
        try:
            import pytesseract
            from PIL import Image
        except ImportError:
            raise ImportError(
                "Tesseract OCR required. Install with:\n"
                "pip install pytesseract pillow\n"
                "Also install Tesseract from: https://github.com/UB-Mannheim/tesseract/wiki"
            )

        image = Image.open(image_path)
        
        # Try to find tesseract.exe if not in PATH
        import shutil
        if not shutil.which("tesseract"):
            common_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                r"C:\Users\User\AppData\Local\Tesseract-OCR\tesseract.exe",
                "/usr/bin/tesseract",
                "/usr/local/bin/tesseract",
            ]
            for path in common_paths:
                if Path(path).exists():
                    pytesseract.pytesseract.tesseract_cmd = path
                    break
        
        text = pytesseract.image_to_string(image)
        return text

    def _extract_text_google_vision(self, image_path: str) -> str:
        """Extract text using Google Cloud Vision API"""
        try:
            from google.cloud import vision
        except ImportError:
            raise ImportError(
                "Google Cloud Vision required. Install with:\n"
                "pip install google-cloud-vision\n"
                "Also set GOOGLE_APPLICATION_CREDENTIALS environment variable"
            )

        client = vision.ImageAnnotatorClient()

        with open(image_path, 'rb') as image_file:
            content = image_file.read()

        image = vision.Image(content=content)
        response = client.document_text_detection(image=image)

        if response.error.message:
            raise Exception(f"Google Vision API Error: {response.error.message}")

        return response.full_text_annotation.text

    def _parse_contribution_text(self, text: str) -> List[ContributionEntry]:
        """
        Parse contribution entries from OCR text

        Expected formats:
        - "John Doe - 5000"
        - "John Doe: 5000 Contribution"
        - "5000 - John Doe (Missions)"
        - "Jane Smith 2500 Benevolence"
        """
        entries = []

        # Split text into lines, strip leading/trailing whitespace and common OCR noise
        # Tesseract often prefixes lines with ' " ` | or similar artifacts
        noise_chars = set("'\"`|\\!@#\u2018\u2019\u201c\u201d\u00b4")
        lines = []
        for line in text.split('\n'):
            line = line.strip()
            # Strip leading noise characters
            while line and line[0] in noise_chars:
                line = line[1:].strip()
            if line:
                lines.append(line)

        for line in lines:
            entry = self._parse_line(line)
            if entry:
                entries.append(entry)

        return entries

    def _parse_line(self, line: str) -> Optional[ContributionEntry]:
        """Parse a single line into a ContributionEntry"""
        # Strip leading OCR noise characters
        noise_chars = set("'\"`|\\!@#\u2018\u2019\u201c\u201d\u00b4")
        while line and line[0] in noise_chars:
            line = line[1:].strip()

        # Skip lines that look like denominations (e.g. "1000 x 5")
        if re.match(r'^\d[\d,]*\s*(?:[xX\*%+\-\/\u00D7]\s*|\s+)\d+', line):
            return None

        # Skip header/label lines with no amount
        skip_labels = {'breakdown', 'contribution cash', 'missions cash',
                       'benevolence cash', 'contribution', 'missions', 'benevolence'}
        if line.lower().strip() in skip_labels:
            return None
        # Pattern 1: Name - Amount [Category]
        pattern1 = r'([A-Za-z\s]+?)\s*[=:-]\s*([\d,]+\.?\d*)\s*(?:\((.+?)\)|(.+?))?$'

        # Pattern 2: Amount - Name [Category]
        pattern2 = r'([\d,]+\.?\d*)\s*[=:-]\s*([A-Za-z\s]+?)(?:\s*\((.+?)\)|\s+(.+?))?$'

        # Pattern 3: Name Amount (simple space separated)
        pattern3 = r'([A-Za-z\s]+?)\s+([\d,]+\.?\d*)\s*(?:\((.+?)\)|\s+(Contribution|Benevolence|Missions))?$'

        for pattern in [pattern1, pattern2, pattern3]:
            match = re.match(pattern, line, re.IGNORECASE)
            if match:
                groups = match.groups()

                # Determine which pattern matched and extract accordingly
                if pattern == pattern1:
                    name = groups[0].strip()
                    amount_str = groups[1]
                    category = (groups[2] or groups[3] or "Contribution").strip()
                elif pattern == pattern2:
                    amount_str = groups[0]
                    name = groups[1].strip()
                    category = (groups[2] or groups[3] or "Contribution").strip()
                else:  # pattern3
                    name = groups[0].strip()
                    amount_str = groups[1]
                    category = (groups[2] or groups[3] or "Contribution").strip()

                # Clean up category
                category = self._normalize_category(category)
                
                # Smart Category detection from Name (Fallback for summary labels)
                if category == "Contribution":
                    name_lower = name.lower()
                    if "miss" in name_lower:
                        category = "Missions"
                    elif "benev" in name_lower:
                        category = "Benevolence"

                # Parse amount
                try:
                    amount = float(amount_str.replace(',', '').replace('KES', '').strip())
                except (ValueError, AttributeError):
                    continue

                # Clean up name
                name = ' '.join(name.split())

                if name and amount > 0:
                    return ContributionEntry(
                        name=name,
                        amount=amount,
                        category=category
                    )

        return None

    def _normalize_category(self, category: str) -> str:
        """Normalize category name"""
        category = category.lower().strip()

        category_map = {
            'contribution': 'Contribution',
            'contrib': 'Contribution',
            'cont': 'Contribution',
            'benevolence': 'Benevolence',
            'benev': 'Benevolence',
            'bene': 'Benevolence',
            'missions': 'Missions',
            'mission': 'Missions',
            'miss': 'Missions'
        }

        return category_map.get(category, 'Contribution')

    def get_total_by_category(self) -> Dict[str, float]:
        """Get total amounts by category"""
        totals = {}
        for entry in self.entries:
            totals[entry.category] = totals.get(entry.category, 0) + entry.amount
        return totals

    def to_dict_list(self) -> List[Dict]:
        """Convert entries to list of dictionaries"""
        return [
            {
                'name': e.name,
                'amount': e.amount,
                'category': e.category,
                'ministry': e.ministry,
                'notes': e.notes
            }
            for e in self.entries
        ]


class ImagePreprocessor:
    """Preprocess images to improve OCR accuracy"""

    @staticmethod
    def enhance_image(image_path: str, output_path: Optional[str] = None) -> str:
        """
        Enhance image for better OCR results

        Args:
            image_path: Path to input image
            output_path: Path to save enhanced image (optional)

        Returns:
            Path to enhanced image
        """
        try:
            from PIL import Image, ImageEnhance, ImageFilter
        except ImportError:
            raise ImportError("Pillow required. Install with: pip install pillow")

        image = Image.open(image_path)

        # Convert to grayscale
        if image.mode != 'L':
            image = image.convert('L')

        # Enhance contrast
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(2.0)

        # Enhance sharpness
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)

        # Apply slight blur to reduce noise
        image = image.filter(ImageFilter.MedianFilter(size=3))

        # Binarize (threshold)
        image = image.point(lambda x: 0 if x < 128 else 255, '1')

        if output_path:
            image.save(output_path)
            return output_path
        else:
            # Save to temp file
            import tempfile
            temp_path = tempfile.mktemp(suffix='.png', prefix='enhanced_')
            image.save(temp_path)
            return temp_path
