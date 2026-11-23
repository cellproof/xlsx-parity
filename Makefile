# One command from a clean checkout: install pinned deps and run the Test 1 baseline.
PYTHON ?= python3
VENV := .venv
PY := $(VENV)/bin/python
PIP := $(VENV)/bin/pip

.PHONY: test
test: $(VENV)/.installed node_modules/.installed
	$(PY) suite/run.py

$(VENV)/.installed: requirements.txt
	$(PYTHON) -m venv $(VENV)
	$(PIP) install --disable-pip-version-check -r requirements.txt
	touch $@

node_modules/.installed: package.json package-lock.json
	npm ci --no-fund --no-audit
	touch $@
