import sys
from utils import get_stats

year  = sys.argv[1] if len(sys.argv) > 1 else None
month = sys.argv[2] if len(sys.argv) > 2 else None

get_stats(year=year, month=month)
