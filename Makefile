PY ?= python3
DEMO := examples/demo.yaml
OUT := examples/output

.PHONY: test eval demo clean report assets briefs prompt-eval

test:
	$(PY) -m unittest discover -s tests -p "test_*.py"

eval:
	$(PY) evals/judgment_eval.py --baseline --ablate --report evals/reports/judgment-baseline.json
	$(PY) evals/typography_eval.py
	$(PY) evals/prompt_eval.py --report evals/reports/prompt-baseline.json

assets:
	$(PY) -m unittest discover -s tests -p "test_assets.py" -v

prompt-eval:
	$(PY) evals/prompt_eval.py --show --report evals/reports/prompt-baseline.json

# 需要生产一张图的最小闭环：plan → build（写出图位 + 生产说明）→ media（导出提示词）
# 之后 check 会 BLOCK / ASSET_MISSING —— 没图就不能宣称完成，这是设计意图
briefs:
	$(PY) scripts/pvad.py plan examples/needs-image.yaml --out $(OUT)/needs-image/plan.json
	$(PY) scripts/pvad.py build $(OUT)/needs-image/plan.json --out $(OUT)/needs-image/deck.json
	$(PY) scripts/pvad.py media $(OUT)/needs-image/deck.json --out $(OUT)/needs-image/media_briefs.json

demo:
	$(PY) scripts/pvad.py plan $(DEMO) --out $(OUT)/plan.json
	$(PY) scripts/pvad.py build $(OUT)/plan.json --brief $(DEMO) --out $(OUT)/deck.json
	$(PY) scripts/pvad.py check $(OUT)/deck.json $(OUT)/demo.pptx --mode release --out-dir $(OUT)/report

report: test eval demo briefs

clean:
	rm -rf $(OUT) evals/reports
	find . -name __pycache__ -type d -prune -exec rm -rf {} +
