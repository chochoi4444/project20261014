"""문장 임베딩 모델로 펀드명을 벡터로 바꾼다 (선택 기능, 인터넷 불필요).

모델은 Hugging Face 형식의 ONNX 파일과 tokenizer.json을 로컬 폴더에 두고 쓴다.
    모델폴더/
      tokenizer.json
      model.onnx            (또는 onnx/model_quantized.onnx, onnx/model.onnx)

필요 라이브러리: onnxruntime, tokenizers, numpy  (기본 기능에는 필요 없음)
"""
from __future__ import annotations

from pathlib import Path

MODEL_CANDIDATES = ["model_quantized.onnx", "onnx/model_quantized.onnx", "model.onnx", "onnx/model.onnx"]


class OnnxEmbedder:
    def __init__(self, model_dir, prefix: str = "", max_len: int = 128):
        import numpy as np
        import onnxruntime as ort
        from tokenizers import Tokenizer

        self.np = np
        d = Path(model_dir)
        path = next((d / c for c in MODEL_CANDIDATES if (d / c).is_file()), None)
        if path is None:
            raise FileNotFoundError(f"{d} 안에서 ONNX 모델({', '.join(MODEL_CANDIDATES)})을 찾지 못했습니다.")
        tok_path = d / "tokenizer.json"
        if not tok_path.is_file():
            raise FileNotFoundError(f"{tok_path} 가 없습니다.")
        self.tok = Tokenizer.from_file(str(tok_path))
        self.tok.enable_truncation(max_length=max_len)
        self.tok.enable_padding(pad_id=self.tok.token_to_id("<pad>") or self.tok.token_to_id("[PAD]") or 0)
        so = ort.SessionOptions()
        so.log_severity_level = 3
        self.sess = ort.InferenceSession(str(path), so, providers=["CPUExecutionProvider"])
        self.inputs = {i.name for i in self.sess.get_inputs()}
        self.prefix = prefix
        self.model_path = path

    def embed(self, texts, batch: int = 32):
        """문장 목록 → (N, D) 단위 벡터. 마지막 은닉층을 attention mask로 평균한 뒤 길이 1로 맞춘다."""
        np = self.np
        out = []
        texts = list(texts)
        for i in range(0, len(texts), batch):
            enc = self.tok.encode_batch([self.prefix + t for t in texts[i:i + batch]])
            ids = np.array([e.ids for e in enc], dtype=np.int64)
            mask = np.array([e.attention_mask for e in enc], dtype=np.int64)
            feed = {"input_ids": ids, "attention_mask": mask}
            if "token_type_ids" in self.inputs:
                feed["token_type_ids"] = np.zeros_like(ids)
            feed = {k: v for k, v in feed.items() if k in self.inputs}
            h = self.sess.run(None, feed)[0]
            if h.ndim == 3:  # (B, T, D) → 평균 풀링
                m = mask[..., None].astype(h.dtype)
                h = (h * m).sum(1) / np.clip(m.sum(1), 1e-9, None)
            out.append(h / np.clip(np.linalg.norm(h, axis=1, keepdims=True), 1e-12, None))
        return np.vstack(out) if out else np.zeros((0, 0))


def cosine_matrix(a, b):
    """단위 벡터끼리의 코사인 유사도 행렬"""
    return a @ b.T
