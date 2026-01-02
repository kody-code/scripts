from datetime import datetime, timedelta

def get_current_date(split: str = "."):
    return datetime.now().strftime(f"%m{split}%d")

def get_tomorrow_date(split: str = "."):
    return (datetime.now() + timedelta(days=1)).strftime(f"%m{split}%d")