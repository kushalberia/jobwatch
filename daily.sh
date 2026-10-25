#!/bin/bash
# jobwatch daily run: sweep every configured board and update tracker.xlsx / pipeline.md.
# Add one line per project folder. Called by launchd (see com.example.jobwatch.plist) or by hand.
cd "$(dirname "$0")" || exit 1
echo "===== $(date '+%Y-%m-%d %H:%M') sweep start"
python3 jobwatch.py 2>&1 | head -60
# second project (optional): ( cd ../agentic && python3 ../jobwatch/jobwatch.py --dir . 2>&1 | head -40 )
echo "===== $(date '+%Y-%m-%d %H:%M') done"
