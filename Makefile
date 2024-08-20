dev.build:
	python -m pip install --upgrade pyinstaller
	pyinstaller --onefile --windowed main.py

dev.clean:
	rm -rf build dist

dev.pushToProd:
	@echo "Pushing to production..."