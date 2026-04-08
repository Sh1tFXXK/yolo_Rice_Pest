from pathlib import Path
import sys

path = Path(sys.argv[1])
patterns = sys.argv[2:]
lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()

for i, line in enumerate(lines, 1):
    if any(p in line for p in patterns):
        start = max(0, i - 3)
        end = min(len(lines), i + 3)
        print(f"=== match line {i} ===")
        for j in range(start, end):
            print(f"{j+1}: {lines[j]}")
