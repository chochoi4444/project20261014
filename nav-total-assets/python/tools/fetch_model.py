"""Hugging Face에서 ONNX 임베딩 모델을 받아 로컬 폴더에 둔다 (인터넷 되는 PC에서 한 번만).
받은 폴더를 사내 PC로 복사하면 인터넷 없이 쓸 수 있다.

    python tools/fetch_model.py e5-small        → models/e5-small/
    python tools/fetch_model.py minilm
    python tools/fetch_model.py <HF저장소> <저장폴더>

모델 후보 (모두 한국어 포함 다국어, ONNX 변환본)
  e5-small  Xenova/multilingual-e5-small                       약 118MB(양자화)  "query: " 접두어 필요
  minilm    Xenova/paraphrase-multilingual-MiniLM-L12-v2       약 118MB(양자화)
  e5-base   Xenova/multilingual-e5-base                        약 280MB(양자화)  "query: " 접두어 필요
"""
import sys
import urllib.request
from pathlib import Path

PRESETS = {
    "e5-small": "Xenova/multilingual-e5-small",
    "minilm": "Xenova/paraphrase-multilingual-MiniLM-L12-v2",
    "e5-base": "Xenova/multilingual-e5-base",
}
FILES = ["tokenizer.json", "config.json", "onnx/model_quantized.onnx"]


def fetch(repo: str, out: Path):
    out.mkdir(parents=True, exist_ok=True)
    for f in FILES:
        dst = out / f
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.exists() and dst.stat().st_size > 0:
            print("있음:", dst)
            continue
        url = f"https://huggingface.co/{repo}/resolve/main/{f}"
        print("받는 중:", url)
        with urllib.request.urlopen(url) as r, open(dst, "wb") as w:
            while chunk := r.read(1 << 20):
                w.write(chunk)
        print(f"  → {dst} ({dst.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    key = sys.argv[1]
    repo = PRESETS.get(key, key)
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).resolve().parent.parent / "models" / (key if key in PRESETS else repo.split("/")[-1])
    fetch(repo, out)
