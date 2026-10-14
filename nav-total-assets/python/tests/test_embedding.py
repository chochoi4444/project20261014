"""임베딩 기능의 배관 점검: 작은 가짜 ONNX 모델로 읽기·풀링·정규화·평가 스크립트가 도는지 본다.
(점수 자체는 의미 없음. onnxruntime·tokenizers·onnx가 없으면 건너뜀)"""
import json
import subprocess
import sys
from pathlib import Path

import pytest

ort = pytest.importorskip("onnxruntime")
pytest.importorskip("tokenizers")
onnx = pytest.importorskip("onnx")
np = pytest.importorskip("numpy")

HERE = Path(__file__).resolve().parent


def make_dummy_model(d: Path, dim=16):
    from onnx import TensorProto, helper, numpy_helper
    from tokenizers import Tokenizer, models, pre_tokenizers, trainers

    d.mkdir(parents=True, exist_ok=True)
    fx = json.loads((HERE.parent.parent / "tests" / "fixtures" / "matching-cases.json").read_text(encoding="utf-8"))
    corpus = fx["targets"] + [v["ledger"] for v in fx["variants"]]
    tok = Tokenizer(models.WordPiece(unk_token="[UNK]"))
    tok.pre_tokenizer = pre_tokenizers.Whitespace()
    tok.train_from_iterator(corpus, trainers.WordPieceTrainer(vocab_size=300, special_tokens=["[PAD]", "[UNK]"]))
    tok.save(str(d / "tokenizer.json"))
    vocab = tok.get_vocab_size()
    table = numpy_helper.from_array(np.random.default_rng(0).normal(size=(vocab, dim)).astype(np.float32), "E")
    g = helper.make_graph(
        [helper.make_node("Gather", ["E", "input_ids"], ["last_hidden_state"])], "dummy",
        [helper.make_tensor_value_info("input_ids", TensorProto.INT64, ["B", "T"]),
         helper.make_tensor_value_info("attention_mask", TensorProto.INT64, ["B", "T"])],
        [helper.make_tensor_value_info("last_hidden_state", TensorProto.FLOAT, ["B", "T", dim])], [table])
    m = helper.make_model(g, opset_imports=[helper.make_opsetid("", 13)])
    m.ir_version = 8
    onnx.save(m, str(d / "model.onnx"))
    return d


def test_embedder_unit_vectors(tmp_path):
    sys.path.insert(0, str(HERE.parent))
    from nav_total_assets.embedding import OnnxEmbedder, cosine_matrix
    e = OnnxEmbedder(make_dummy_model(tmp_path / "m"))
    v = e.embed(["가온대체투자 101호", "누리 해외부동산 1호", "가온대체투자 101호"])
    assert v.shape == (3, 16)
    assert np.allclose(np.linalg.norm(v, axis=1), 1)
    c = cosine_matrix(v, v)
    assert c[0, 2] == pytest.approx(1, abs=1e-6)


def test_eval_script_runs(tmp_path):
    d = make_dummy_model(tmp_path / "m")
    r = subprocess.run([sys.executable, str(HERE.parent / "tools" / "embedding_eval.py"), "--model", str(d)],
                       capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    assert "현재" in r.stdout and "임베딩+규칙" in r.stdout
