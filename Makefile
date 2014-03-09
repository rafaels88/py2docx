clean:
	@echo "Cleaning up build and *.pyc files..."
	@find . -name '*.pyc' -exec rm -rf {} \;
	@rm -rf build

requirements:
	@echo "Installing DEV requirements"
	@pip install -r requirements_dev.txt

unit: clean
	@echo "Running Unit Tests"
	@nosetests -s --verbose --with-coverage \
		--cover-package=py2docx tests/unit/*
