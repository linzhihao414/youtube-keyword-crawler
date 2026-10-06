"""
YouTube Keyword Crawler
Search keywords -> collect video info, channel info, stats -> export Excel
Edit keywords.txt to change search terms.
"""
import asyncio
import re
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

# ============ Config ============
KEYWORDS_FILE = Path(__file__).parent / "keywords.txt"
MAX_RESULTS_PER_KEYWORD = 50
OUTPUT_DIR = Path(__file__).parent / "output"
# ================================


def load_keywords():
    if KEYWORDS_FILE.exists():
        text = KEYWORDS_FILE.read_text(encoding="utf-8").strip()
        return [k.strip() for k in re.split(r"[,，\n]", text) if k.strip()]
    return ["leather bag"]


def search_videos(keyword, max_results=50):
    """Search YouTube and return video info using yt-dlp"""
    import yt_dlp

    ydl_opts = {
        "quiet": True,
        "no_warnings": True,
        "extract_flat": True,
        "skip_download": True,
        "default_search": f"ytsearch{max_results}",
    }

    results = []
    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(f"ytsearch{max_results}:{keyword}", download=False)
        for entry in info.get("entries", []):
            if not entry:
                continue
            results.append({
                "keyword": keyword,
                "title": entry.get("title", ""),
                "video_id": entry.get("id", ""),
                "url": f"https://www.youtube.com/watch?v={entry.get('id','')}",
                "channel": entry.get("channel", entry.get("uploader", "")),
                "channel_id": entry.get("channel_id", entry.get("uploader_id", "")),
                "channel_url": f"https://www.youtube.com/channel/{entry.get('channel_id','')}" if entry.get("channel_id") else "",
                "duration": entry.get("duration", 0),
                "view_count": entry.get("view_count", 0),
                "like_count": entry.get("like_count", 0),
                "comment_count": entry.get("comment_count", 0),
                "upload_date": entry.get("upload_date", ""),
                "description": (entry.get("description", "") or "")[:500],
                "collect_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
            })
    return results


def save_to_excel(all_data, keyword):
    OUTPUT_DIR.mkdir(exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "YouTube Results"

    headers = ["Keyword", "Title", "URL", "Channel", "Channel URL",
               "Views", "Likes", "Comments", "Duration(s)",
               "Upload Date", "Description", "Collected At"]
    ws.append(headers)

    for v in all_data:
        ws.append([
            v["keyword"], v["title"], v["url"], v["channel"], v["channel_url"],
            v["view_count"], v["like_count"], v["comment_count"], v["duration"],
            v["upload_date"], v["description"], v["collect_time"]
        ])

    # Column widths
    widths = [15, 40, 45, 25, 45, 10, 8, 10, 10, 12, 50, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w

    fname = OUTPUT_DIR / f"YouTube_{keyword}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb.save(str(fname))
    print(f"  Saved: {fname}")
    return str(fname)


def main():
    keywords = load_keywords()
    print("=" * 50)
    print("  YouTube Keyword Crawler")
    print("=" * 50)
    print(f"Keywords: {', '.join(keywords)}")
    print(f"Max {MAX_RESULTS_PER_KEYWORD} videos per keyword")
    print()

    total = 0
    for kw in keywords:
        print(f"\nSearching: {kw}")
        try:
            videos = search_videos(kw, MAX_RESULTS_PER_KEYWORD)
            print(f"  Found {len(videos)} videos")
            save_to_excel(videos, kw)
            total += len(videos)
        except Exception as e:
            print(f"  Error: {e}")

    print(f"\nDone! Total {total} videos collected.")
    print(f"Output: {OUTPUT_DIR.resolve()}")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
