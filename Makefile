install:
	python -m venv .env
	.env\Scripts\activate
	pip install -r requirements.txt

run:
	cd test && pytest -s