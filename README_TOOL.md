Local Ollama document->spreadsheet tool

Overview
- This repository contains a small Python CLI `ollama_tool.py` that:
  - Extracts text from PDF, DOCX, or DOC (via LibreOffice conversion)
  - Builds a prompt combining your instructions and the document text
  - Automatically chunks large documents and processes them in parallel
  - Runs a local Ollama command (configurable) to produce structured JSON
  - Saves structured results to CSV or XLSX

Quick setup
1. Create a Python venv and install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Install LibreOffice if you need to process `.doc` files (used for conversion to `.docx`).
3. Install and run Ollama locally and make sure you can call it from the command line.

Usage

Basic example (you must provide a prompt describing the fields to extract):

```bash
python ollama_tool.py path/to/doc.pdf --prompt "Extract Title, Author, Date, and all bullet items as 'items'" --model my-local-model --out results.xlsx
```

Using a prompt file:

```bash
python ollama_tool.py invoice.pdf --prompt "$(cat example_prompt_invoice.txt)" --model llama2 --out invoices.csv
```

Chunking for large documents
By default, documents larger than 8000 characters are automatically split into overlapping chunks (500 char overlap). Each chunk is processed independently and results are merged. You can customize the chunk size:

```bash
python ollama_tool.py large_document.pdf --prompt "..." --chunk-size 12000 --model llama2
```

OLLAMA command customization
- The script runs an external command template to invoke Ollama. Set the `OLLAMA_CMD` environment variable to match your local invocation. The template should include `{model}` and `{prompt_file}` placeholders.
- Default template: `ollama run {model} --prompt-file {prompt_file}`

Example if your ollama CLI expects a different form:

```bash
export OLLAMA_CMD="ollama run {model} --prompt-file {prompt_file} --temperature 0.2"
```

Example prompts
See the included `.txt` files:
- `example_prompt_invoice.txt` — Extract invoice data (number, date, vendor, amount, etc.)
- `example_prompt_research.txt` — Extract research paper metadata (title, authors, abstract, etc.)

You can use these as templates and modify them for your own extraction tasks.

Notes
- The tool expects the model's output to be valid JSON (either a list of objects). Chunks are merged into a single flat list.
- For `.doc` files, the script uses LibreOffice `soffice` to convert to `.docx`.
- If all chunks fail to parse JSON, the tool prints a warning. Consider refining your prompt if this happens.

## How it works

The tool follows a simple pipeline:

1. **Text Extraction** — Reads the input document (PDF/DOCX/DOC) and extracts all text.
2. **Chunking** — If the document exceeds the chunk size threshold (default 8000 chars), it's split into overlapping chunks.
3. **Prompt Building** — For each chunk, a full prompt is constructed containing system instructions, your custom extraction prompt, and the document text.
4. **Model Invocation** — Runs the Ollama model on each chunk via a configurable shell command.
5. **JSON Parsing** — Extracts JSON from the model's output (tolerant parser handles extra text).
6. **Merging & Export** — Combines results from all chunks into a single list and saves to CSV or XLSX.

## Code structure

**[ollama_tool.py](ollama_tool.py)** — Main module with:

| Function | Purpose |
|----------|---------|
| `extract_pdf(path)` | Uses PyPDF2 to read PDF and extract text from all pages |
| `extract_docx(path)` | Uses python-docx to read DOCX and extract paragraph text |
| `convert_doc_to_docx(doc_path)` | Uses LibreOffice to convert legacy `.doc` to `.docx` format |
| `extract_text(file_path)` | Dispatcher that calls the appropriate extractor based on file extension |
| `chunk_text(text, chunk_size, overlap)` | Splits text into overlapping chunks for large documents |
| `build_prompt(user_prompt, doc_text)` | Combines system instructions, your prompt, and document text |
| `run_ollama_with_cmd(model, prompt_file)` | Executes Ollama via shell command, returns model output |
| `try_parse_json(s)` | Tolerant JSON parser—finds JSON objects/arrays within text |
| `save_structured(output_obj, out_path)` | Converts parsed JSON to pandas DataFrame and exports to CSV/XLSX |
| `main()` | CLI entry point; orchestrates the full pipeline |

**[example_prompt_invoice.txt](example_prompt_invoice.txt)** — Sample extraction prompt for invoice data.

**[example_prompt_research.txt](example_prompt_research.txt)** — Sample extraction prompt for research paper metadata.

**[requirements.txt](requirements.txt)** — Python dependencies:
- `PyPDF2` — PDF text extraction
- `python-docx` — DOCX text extraction
- `pandas` — DataFrame creation and CSV/XLSX export
- `openpyxl` — XLSX file support

## Customization points

- **Prompt engineering:** Edit or create new `.txt` prompt files to define what fields to extract.
- **Ollama command:** Set `OLLAMA_CMD` env var to customize how Ollama is invoked (e.g., add temperature or other parameters).
- **Chunk size:** Use `--chunk-size` flag to adjust splitting behavior for your document types.
- **Output format:** Use `--out` with `.csv` or `.xlsx` extension to control output format.

Next steps / improvements you can ask me to implement
- Add automatic prompt optimization based on document structure
- Support CSV output formatting options (delimiter, encoding)
- Add filtering and post-processing of extracted records
- Batch process multiple files in a directory


