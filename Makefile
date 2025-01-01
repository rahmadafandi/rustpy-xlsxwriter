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
	pytest --durations=0 -v

clean:
	rm -rf target
	rm -rf .pytest_cache
	rm -rf tmp/test_*.xlsx