# -*- coding: utf-8 -*-
"""微调模型导出 ONNX：scripts/models/bge-small-zh-v1.5-ft -> model.onnx（fp32）。
用 .ftvenv 运行（需 torch）；量化在主环境用 onnxruntime.quantization 单独跑。
跑法：.ftvenv\\Scripts\\python.exe scripts\\export_ft_onnx.py
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FT_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    BASE, "scripts", "models", "bge-small-zh-v1.5-ft")
if not os.path.isabs(FT_DIR):
    FT_DIR = os.path.join(BASE, FT_DIR)
OUT_PATH = os.path.join(FT_DIR, "model.onnx")
MAXLEN = 48


def main():
    import torch
    from sentence_transformers import SentenceTransformer, models

    word = models.Transformer(FT_DIR, max_seq_length=MAXLEN)
    pool = models.Pooling(word.get_word_embedding_dimension(),
                          pooling_mode_cls_token=True)
    model = SentenceTransformer(modules=[word, pool])
    model.eval()

    tok = word.tokenizer
    enc = tok(["沉头钻尾螺钉 十字沉头螺丝"], return_tensors="pt",
              padding="max_length", truncation=True, max_length=MAXLEN)
    inputs = {k: v for k, v in enc.items() if k in ("input_ids", "attention_mask", "token_type_ids")}

    class Wrap(torch.nn.Module):
        def __init__(self, m):
            super().__init__()
            self.m = m

        def forward(self, input_ids, attention_mask, token_type_ids=None):
            feats = {"input_ids": input_ids, "attention_mask": attention_mask}
            if token_type_ids is not None:
                feats["token_type_ids"] = token_type_ids
            return self.m(feats)["token_embeddings"]

    wrap = Wrap(model)
    wrap.eval()
    with torch.no_grad():
        torch.onnx.export(
            wrap,
            (inputs["input_ids"], inputs["attention_mask"],
             inputs.get("token_type_ids")),
            OUT_PATH,
            input_names=["input_ids", "attention_mask", "token_type_ids"],
            output_names=["embedding"],
            dynamic_axes={n: {0: "batch", 1: "seq"} for n in
                          ["input_ids", "attention_mask", "token_type_ids", "embedding"]},
            opset_version=14,
            do_constant_folding=True,
            dynamo=False)
    print("exported:", OUT_PATH, os.path.getsize(OUT_PATH) // 1024, "KB")

    import onnx
    m = onnx.load(OUT_PATH)
    onnx.checker.check_model(m)
    print("onnx check ok, inputs:", [i.name for i in m.graph.input])


if __name__ == "__main__":
    main()
