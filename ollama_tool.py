#!/usr/bin/env python3
import argparse
import os
import shlex
import shutil
import subprocess
import sys
import tempfile
import json
from pathlib import Path

import pandas as pd

try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

try:
    from docx import Document
except Exception:
    Document = None


def extract_pdf(path: Path) -> str:
    if PdfReader is None:
        raise RuntimeError("PyPDF2 not installed. Install from requirements.txt")
    reader = PdfReader(str(path))
    pages = []
    for p in reader.pages:
        pages.append(p.extract_text() or "")
    return "\n\n".join(pages)


def extract_docx(path: Path) -> str:
    if Document is None:
        raise RuntimeError("python-docx not installed. Install from requirements.txt")
    doc = Document(str(path))
    paragraphs = [p.text for p in doc.paragraphs]
    return "\n\n".join(paragraphs)


def convert_doc_to_docx(doc_path: Path) -> Path:
    outdir = Path(tempfile.mkdtemp(prefix="doc2docx_"))
    soffice = shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice 'soffice' not found. Install LibreOffice to convert .doc to .docx")
    cmd = [soffice, "--headless", "--convert-to", "docx", str(doc_path), "--outdir", str(outdir)]
    subprocess.run(cmd, check=True)
    converted = list(outdir.glob("*.docx"))
    if not converted:
        raise RuntimeError("Conversion to .docx failed")
    return converted[0]


def extract_text(file_path: str) -> str:
    path = Path(file_path)
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(path)
    if suffix == ".docx":
        return extract_docx(path)
    if suffix == ".doc":
        converted = convert_doc_to_docx(path)
        text = extract_docx(converted)
        return text
    raise RuntimeError(f"Unsupported file extension: {suffix}")


def chunk_text(text: str, chunk_size: int = 8000, overlap: int = 500) -> list:
    """Split text into overlapping chunks for processing large documents."""
    chunks = []
    start = 0
    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunks.append(text[start:end])
        start = end - overlap
        if start >= len(text):
            break
    return chunks if chunks else [text]


def build_prompt(user_prompt: str, doc_text: str) -> str:
    header = (
        "You are a tool that extracts structured data from the provided document.\n"
        "Return the result in strict JSON format, suitable for loading into a spreadsheet (a list of objects).\n"
        "Document follows below. Use the user's prompt to decide what fields to extract.\n\n"
    )
    return header + "USER INSTRUCTIONS:\n" + user_prompt + "\n\nDOCUMENT TEXT:\n" + doc_text


def run_ollama_with_cmd(model: str, prompt_file: str) -> str:
    cmd_template = os.environ.get(
        "OLLAMA_CMD",
        "ollama run {model} --prompt-file {prompt_file}"
    )
    cmd = cmd_template.format(model=shlex.quote(model), prompt_file=shlex.quote(prompt_file))
    proc = subprocess.run(cmd if isinstance(cmd, str) else cmd, shell=True, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"Ollama command failed: {proc.stderr}\nCommand: {cmd}")
    return proc.stdout


def try_parse_json(s: str):
    try:
        return json.loads(s)
    except Exception:
        try:
            start = s.index("{")
            end = s.rindex("}")
            return json.loads(s[start:end+1])
        except Exception:
            try:
                start = s.index("[")
                end = s.rindex("]")
                return json.loads(s[start:end+1])
            except Exception:
                return None


def save_structured(output_obj, out_path: Path):
    if isinstance(output_obj, list):
        df = pd.DataFrame(output_obj)
    elif isinstance(output_obj, dict):
        if all(isinstance(v, list) for v in output_obj.values()):
            df = pd.DataFrame(output_obj)
        else:
            df = pd.DataFrame([output_obj])
    else:
        raise RuntimeError("Parsed output is not list/dict, cannot convert to table")
    out_path_parent = out_path.parent
    out_path_parent.mkdir(parents=True, exist_ok=True)
    if out_path.suffix.lower() in [".xlsx", ".xls"]:
        df.to_excel(out_path, index=False)
    else:
        df.to_csv(out_path, index=False)


def main():
    parser = argparse.ArgumentParser(description="Extract text from a document and run a local Ollama model to produce structured output")
    parser.add_argument("input", help="Input file (PDF, DOC, DOCX)")
    parser.add_argument("--model", default="llama2", help="Local Ollama model name")
    parser.add_argument("--prompt", required=True, help="User prompt describing the structure to extract")
    parser.add_argument("--out", default="output.csv", help="Output CSV or XLSX path")
    parser.add_argument("--chunk-size", type=int, default=8000, help="Text chunk size for large documents (chars)")
    parser.add_argument("--temp-prompt", default=None, help="(optional) path to write the full prompt file")
    args = parser.parse_args()

    doc_text = extract_text(args.input)
    
    # Split into chunks if document is large
    chunks = chunk_text(doc_text, chunk_size=args.chunk_size)
    all_parsed = []
    
    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            print(f"Processing chunk {i+1}/{len(chunks)}...")
        
        full_prompt = build_prompt(args.prompt, chunk)
        if args.temp_prompt:
            with open(args.temp_prompt, "w") as f:
                f.write(full_prompt)
            prompt_file = args.temp_prompt
        else:
            fd, tmp = tempfile.mkstemp(prefix="ollama_prompt_", suffix=".txt")
            with os.fdopen(fd, "w") as f:
                f.write(full_prompt)
            prompt_file = tmp

        try:
            out = run_ollama_with_cmd(args.model, prompt_file)
        finally:
            if not args.temp_prompt:
                try:
                    os.remove(prompt_file)
                except Exception:
                    pass

        parsed = try_parse_json(out)
        if parsed is not None:
            if isinstance(parsed, list):
                all_parsed.extend(parsed)
            elif isinstance(parsed, dict):
                all_parsed.append(parsed)
    
    out_path = Path(args.out)
    if all_parsed:
        save_structured(all_parsed, out_path)
        print(f"Saved structured output to {out_path} ({len(all_parsed)} records)")
    else:
        print("Warning: No valid JSON parsed from model output. Check the prompt and model.")


if __name__ == "__main__":
    main()
