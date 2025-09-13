from datetime import datetime, date
from typing import Union, Optional
import locale
import re

def format_currency(amount: Union[int, float]) -> str:
    """Format currency values in Arabic locale"""
    try:
        if amount is None:
            return "0 ريال"
        return f"{amount:,.2f} ريال"
    except:
        return "0 ريال"

def format_percentage(value: Union[int, float], decimal_places: int = 1) -> str:
    """Format percentage values"""
    try:
        if value is None:
            return "0%"
        return f"{value:.{decimal_places}f}%"
    except:
        return "0%"

def validate_date_range(start_date: date, end_date: date) -> bool:
    """Validate that end date is after start date"""
    try:
        return end_date >= start_date
    except:
        return False

def calculate_days_difference(start_date: date, end_date: date) -> int:
    """Calculate the number of days between two dates"""
    try:
        return (end_date - start_date).days
    except:
        return 0

def parse_arabic_number(text: str) -> Optional[float]:
    """Parse Arabic number text to float"""
    try:
        # Replace Arabic digits with English digits
        arabic_to_english = {
            '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
            '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9'
        }
        
        english_text = text
        for arabic, english in arabic_to_english.items():
            english_text = english_text.replace(arabic, english)
        
        # Remove any non-numeric characters except decimal point and minus sign
        english_text = re.sub(r'[^\d.-]', '', english_text)
        
        return float(english_text) if english_text else None
    except:
        return None

def format_date_arabic(date_obj: Union[date, datetime, str]) -> str:
    """Format date in Arabic style"""
    try:
        if isinstance(date_obj, str):
            date_obj = datetime.strptime(date_obj, '%Y-%m-%d').date()
        elif isinstance(date_obj, datetime):
            date_obj = date_obj.date()
        
        # Arabic month names
        arabic_months = {
            1: 'يناير', 2: 'فبراير', 3: 'مارس', 4: 'أبريل',
            5: 'مايو', 6: 'يونيو', 7: 'يوليو', 8: 'أغسطس',
            9: 'سبتمبر', 10: 'أكتوبر', 11: 'نوفمبر', 12: 'ديسمبر'
        }
        
        day = date_obj.day
        month = arabic_months[date_obj.month]
        year = date_obj.year
        
        return f"{day} {month} {year}"
    except:
        return str(date_obj)

def validate_project_name(name: str) -> bool:
    """Validate project name"""
    try:
        if not name or len(name.strip()) == 0:
            return False
        if len(name.strip()) > 200:
            return False
        # Check for invalid characters (optional)
        invalid_chars = ['<', '>', ':', '"', '|', '?', '*']
        for char in invalid_chars:
            if char in name:
                return False
        return True
    except:
        return False

def clean_text_input(text: str) -> str:
    """Clean and sanitize text input"""
    try:
        if not text:
            return ""
        # Remove extra whitespace
        cleaned = re.sub(r'\s+', ' ', text.strip())
        # Remove potentially harmful characters
        cleaned = re.sub(r'[<>]', '', cleaned)
        return cleaned
    except:
        return ""

def calculate_project_duration(start_date: date, end_date: date) -> dict:
    """Calculate project duration in various units"""
    try:
        delta = end_date - start_date
        days = delta.days
        weeks = days // 7
        months = days // 30  # Approximate
        years = days // 365  # Approximate
        
        return {
            'days': days,
            'weeks': weeks,
            'months': months,
            'years': years
        }
    except:
        return {'days': 0, 'weeks': 0, 'months': 0, 'years': 0}

def get_project_status_color(status: str) -> str:
    """Get color code for project status"""
    status_colors = {
        'متقدم': '#28a745',      # Green
        'على المسار': '#17a2b8',   # Blue
        'متأخر': '#dc3545',      # Red
        'مكتمل': '#6f42c1',      # Purple
        'متوقف': '#6c757d'       # Gray
    }
    return status_colors.get(status, '#6c757d')

def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format"""
    try:
        if size_bytes == 0:
            return "0 بايت"
        
        size_names = ["بايت", "كيلوبايت", "ميجابايت", "جيجابايت"]
        i = 0
        size = float(size_bytes)
        
        while size >= 1024.0 and i < len(size_names) - 1:
            size /= 1024.0
            i += 1
        
        return f"{size:.1f} {size_names[i]}"
    except:
        return "0 بايت"

def get_completion_status(planned_completion: float, actual_completion: float) -> str:
    """Get completion status based on planned vs actual"""
    try:
        if actual_completion >= planned_completion:
            return "متقدم عن الخطة"
        elif actual_completion >= planned_completion * 0.9:
            return "ضمن الخطة"
        else:
            return "متأخر عن الخطة"
    except:
        return "غير محدد"

def validate_budget_input(budget: Union[str, int, float]) -> bool:
    """Validate budget input"""
    try:
        if isinstance(budget, str):
            budget = parse_arabic_number(budget)
        
        if budget is None or budget <= 0:
            return False
        
        # Check if budget is reasonable (not too large)
        if budget > 1e12:  # 1 trillion
            return False
        
        return True
    except:
        return False

def calculate_burn_rate(actual_costs: list, dates: list) -> float:
    """Calculate average burn rate (cost per day)"""
    try:
        if len(actual_costs) < 2 or len(dates) < 2:
            return 0.0
        
        total_cost = actual_costs[-1] - actual_costs[0]
        start_date = datetime.strptime(dates[0], '%Y-%m-%d') if isinstance(dates[0], str) else dates[0]
        end_date = datetime.strptime(dates[-1], '%Y-%m-%d') if isinstance(dates[-1], str) else dates[-1]
        
        days = (end_date - start_date).days
        
        if days <= 0:
            return 0.0
        
        return total_cost / days
    except:
        return 0.0

def get_arabic_weekday(date_obj: Union[date, datetime]) -> str:
    """Get Arabic weekday name"""
    try:
        if isinstance(date_obj, datetime):
            date_obj = date_obj.date()
        
        arabic_weekdays = {
            0: 'الاثنين',
            1: 'الثلاثاء', 
            2: 'الأربعاء',
            3: 'الخميس',
            4: 'الجمعة',
            5: 'السبت',
            6: 'الأحد'
        }
        
        return arabic_weekdays[date_obj.weekday()]
    except:
        return "غير محدد"

def export_data_summary(data: dict) -> str:
    """Create a text summary of data for export"""
    try:
        summary_lines = []
        summary_lines.append("ملخص بيانات المشروع")
        summary_lines.append("=" * 50)
        
        for key, value in data.items():
            if isinstance(value, (int, float)):
                if 'تكلفة' in key or 'ميزانية' in key:
                    summary_lines.append(f"{key}: {format_currency(value)}")
                elif 'نسبة' in key or 'مؤشر' in key:
                    summary_lines.append(f"{key}: {format_percentage(value)}")
                else:
                    summary_lines.append(f"{key}: {value}")
            else:
                summary_lines.append(f"{key}: {value}")
        
        return "\n".join(summary_lines)
    except:
        return "خطأ في إنشاء الملخص"
