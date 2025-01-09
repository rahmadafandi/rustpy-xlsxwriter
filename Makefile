build-development:
	maturin develop

build-development-release:
	maturin develop --release

build-production:
	maturin build --release

build-wheel:
	maturin build --release --wheel

publish:
	maturin publish

install-development:
	pip install .

install-production:
	pip install dist/excel_exporter-*.whl

test:
	pytest tests/ -v -ra --durations=0 -o log_cli=true

test-benchmark:
	pytest tests/ -v -ra --durations=0 -o log_cli=true --codspeed

clean:
	rm -rf target
	rm -rf .pytest_cache
	rm -rf tmp/test_*.xlsx

cleanup:
	autoflake --remove-unused-variables --remove-all-unused-imports -i --recursive . && black . && isort --profile black . && pyclean .

requirements:
	pip freeze > requirements.txt
