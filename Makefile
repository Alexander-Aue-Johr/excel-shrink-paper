
.PHONY: run venv activate clean pdf

# Create virtual environment if it does not exist
venv:
	@if [ ! -d ".venv" ]; then \
		python3 -m venv .venv; \
		echo "✅ Virtual environment created."; \
	else \
		echo "✅ Virtual environment already exists."; \
	fi

# Activate venv and run the main benchmark script
run: venv
	@.venv/bin/python scripts/loading_times_size_reduction/loading_times_size_reduction_chart.py --input-folder input --shrunk-folder shrunk

# Clean up virtual environment and output files
clean:
	@rm -rf .venv
	@rm -f excel_benchmarks.csv
	@rm -f time_and_filesize_comparison_by_file.pdf
	@rm -rf __pycache__
	@echo "🧹 Clean complete."

# Open the output PDF (Linux/macOS only)
pdf:
	@xdg-open time_and_filesize_comparison_by_file.pdf || open time_and_filesize_comparison_by_file.pdf || echo "📄 Please open the PDF manually."
