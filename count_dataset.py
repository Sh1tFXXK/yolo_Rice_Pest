from pathlib import Path

base = Path(r"E:\code\yolo_Rice_Pest\train_data")
for split in ["train", "valid", "test"]:
    img_dir = base / split / "images"
    count = sum(1 for p in img_dir.iterdir() if p.is_file())
    print(split, count)
