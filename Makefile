dev.build:
	@echo "running updates..."
	python -m pip install --upgrade pyinstaller
	@echo "Running pyinstaller..."
	pyinstaller --onefile main.py
	@echo "completed!"

dev.clean:
	rm -rf build dist

dev.pushToProd:
	@echo "Pushing to production..."