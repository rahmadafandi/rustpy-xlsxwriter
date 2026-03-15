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
	pip install dist/rustpy_xlsxwriter-*.whl

test:
	pytest tests/ -v -ra --durations=0 -m "not benchmark"

test-all:
	pytest tests/ -v -ra --durations=0

test-benchmark:
	pytest tests/test_benchmark.py -v -ra --durations=0 --codspeed

clean:
	rm -rf target
	rm -rf .pytest_cache

cleanup:
	autoflake --remove-unused-variables --remove-all-unused-imports -i --recursive . && black . && isort --profile black . && pyclean .

requirements:
	pip freeze > requirements.txt
