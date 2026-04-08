from pathlib import Path
import sys

patterns = sys.argv[1:-1]
path = Path(sys.argv[-1])

text = path.read_text(encoding="utf-8", errors="ignore").splitlines()
for i, line in enumerate(text, 1):
    if any(p.lower() in line.lower() for p in patterns):
        sys.stdout.buffer.write(f"{path.name}:{i}:{line}\n".encode("utf-8", errors="replace"))
