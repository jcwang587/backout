import sys
from utils import get_stats

args = sys.argv[1:]
show_top_senders = "--top-senders" in args
args = [arg for arg in args if arg != "--top-senders"]

year  = args[0] if len(args) > 0 else None
month = args[1] if len(args) > 1 else None

get_stats(year=year, month=month, show_top_senders=show_top_senders)
