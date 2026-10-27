.PHONY: test lint generate-sample

test:
	pytest -q

lint:
	ruff check src tests
	ruff format --check src tests
	mypy src

# Generate the fictional sample workbook into ./output (git-ignored).
generate-sample:
	tech6-team-composition generate examples/validation.json
