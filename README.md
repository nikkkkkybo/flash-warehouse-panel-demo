# 闪电仓选品 / 竞对对齐面板（公开演示版）

即时零售小店（五金、日百这类闪电仓）用的本地工具：把自家商品库和竞店导出表放进同一个口径对齐，
再在浏览器面板里看竞对、找价差、选品、出 AI 建议。单机运行，不用服务器也不用数据库——数据都是本地文件。

Python 3.11 + Flask + pandas + waitress；AI 侧走 DeepSeek 兼容接口，可以完全关掉。

## 这份和内部版差在哪

这里是**脱敏镜像**，只去掉/替换了三类东西，逻辑是同一套：

- **不含内部文档**：方案文档、任务规划、逐轮施工日志留在私有仓库。
- **不含自检与一次性脚本**：`scripts/_*.py`（`_verify_*` 回归、`_package.py` 打包、训练探针等）没发出来——
  它们的断言写死了真实店名和真实数据，去掉店名就失去意义。
- **不含真实身份**：店名、美团店/仓编码、导入映射都换成了 `示例竞店A~E`、`示例自家店1~3`、`1000001` 这类占位；
  `config/analysis_shops.yaml`（选店清单）和 `config/store_warehouse_map.yaml`（店↔仓映射）没发——
  这两个文件是面板自己读写生成的，你导入自己的数据就会长出你自己的那份。

模型权重（BGE-small-zh-v1.5 的 ONNX，约 700MB）同样不在仓库里。缺失时语义/向量层自动跳过，其余照常。

## 跑起来

```bash
pip install -r requirements.txt
cp config/ai_config.example.yaml config/ai_config.yaml   # 要用 AI 才填 key；不填就是 driver: fake，不花钱
python scripts/launch_panel.py                            # 自动打开 http://127.0.0.1:5000
```

面板本身不带数据。要看到东西：按 `自家商品库_标准模板.xlsx` 填好自家库放在根目录，
竞店导出表在「导入」页上传并选一个列映射（`config/import_profiles/`），归一化后落进 `output/`。
之后的对齐结果、快照缓存、价格观测都由面板重算生成——每个派生文件绑"源文件指纹 + 生成代码指纹"，代码一改缓存自动作废。

## 核心口径：销量文本

竞店表的「月售」一格里混着两种东西：`月售35` 是**真成交**，`35人想买` 是**收藏/意向**。
不能等量齐观，但意向也不是没意义——它是有信号的弱证据。全站内置同一个公式：

```
有效销量 = 成交 + λ × 收藏        （λ = 0.2，环境变量 INTENT_WEIGHT 可调）
```

解析只在 `scripts/sales_text.py` 一处定义，探测器、AI 建议、抽屉、排序、人工池过滤都从这里取数；
提示词里的 λ 由代码注入，不会出现"文字说一套、代码算一套"。

## 目录

- `scripts/web_panel.py` 面板与路由（101 条）；`scripts/align.py` 四层对齐引擎：条码 → 规则 → 语义 → 人工池。
- `scripts/snapcache.py` 快照缓存层；`scripts/insight_scan.py` + `insight_store.py` 机会探测器与台账。
- `scripts/assistant_brain.py` + `scripts/templates/_assistant.html` 页面右侧 AI 小助手抽屉。
- `config/` 列映射、同义词、类目树、AI 配置模板；`scripts/templates/` 页面，`scripts/static/` 样式。
