import os
import re
import hashlib
from pathlib import Path
from typing import List, Dict, Tuple

import pdfplumber
from docx import Document
from pptx import Presentation
import extract_msg
import pandas as pd

from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

import openai
from dotenv import load_dotenv

load_dotenv()
openai.api_type = "azure"
openai.api_key = os.getenv("AZURE_OPENAI_KEY")
openai.api_base = os.getenv("AZURE_OPENAI_ENDPOINT")
openai.api_version = os.getenv("AZURE_OPENAI_VERSION", "2023-05-15")
MODEL_NAME = os.getenv("AZURE_OPENAI_DEPLOYMENT")  # deployment name


def extract_text_from_file(file_path: Path) -> str:
    """Extract visible text from supported file types."""
    ext = file_path.suffix.lower()
    if ext == ".pdf":
        text = ""
        with pdfplumber.open(str(file_path)) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""
        return text
    elif ext in {".docx", ".doc"}:
        doc = Document(str(file_path))
        return "\n".join(p.text for p in doc.paragraphs)
    elif ext in {".pptx", ".ppt"}:
        prs = Presentation(str(file_path))
        slides_text = []
        for slide in prs.slides:
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_text.append(shape.text)
            slides_text.append("\n".join(slide_text))
        return "\n".join(slides_text)
    elif ext == ".msg":
        msg = extract_msg.Message(str(file_path))
        parts = [msg.subject or "", msg.body or ""]
        return "\n".join(parts)
    else:
        raise ValueError(f"Unsupported file type: {ext}")


def call_llm(prompt: str, temperature: float = 0.1, max_tokens: int = 512) -> str:
    """Call Azure OpenAI chat completion."""
    response = openai.ChatCompletion.create(
        engine=MODEL_NAME,
        messages=[{"role": "system", "content": prompt}],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    return response["choices"][0]["message"]["content"].strip()


def summarize_text(text: str) -> Tuple[str, str, str, str]:
    """Get summary, version, tags from LLM."""
    system_prompt = (
        "You are an assistant that summarizes documents. "
        "Provide a concise paragraph covering all necessary content, followed by 2-3 bullet points. "
        "After the summary, output on new lines in the format: \nFILE_TYPE: <type>\nVERSION: <version>\nTAGS: <comma separated keywords>."""
    )
    combined_prompt = f"""{system_prompt}\n\nDOCUMENT:\n{text[:8000]}"""

    output = call_llm(combined_prompt)

    # Regex parse
    summary_part, file_type, version, tags = "", "", "", ""
    try:
        lines = output.splitlines()
        # find markers
        file_type_line = next(l for l in lines if l.startswith("FILE_TYPE:"))
        version_line = next(l for l in lines if l.startswith("VERSION:"))
        tags_line = next(l for l in lines if l.startswith("TAGS:"))
        idx_ft = lines.index(file_type_line)
        summary_part = "\n".join(lines[:idx_ft]).strip()
        file_type = file_type_line.split(":", 1)[1].strip()
        version = version_line.split(":", 1)[1].strip()
        tags = tags_line.split(":", 1)[1].strip()
    except Exception:
        summary_part = output.strip()
    return summary_part, file_type, version, tags


def compute_hash(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def detect_duplicates(rows: List[Dict]) -> None:
    """Add duplicate info in-place among rows list."""
    model = SentenceTransformer("all-MiniLM-L6-v2")
    embeddings = model.encode([r["extracted_text"] for r in rows], convert_to_tensor=True)
    sim_matrix = cosine_similarity(embeddings, embeddings)
    n = len(rows)
    master_indices = {}
    for i in range(n):
        for j in range(i + 1, n):
            score = sim_matrix[i, j]
            if score > 0.9:  # threshold for duplicate
                # decide master as longer text version
                master = i if len(rows[i]["extracted_text"]) >= len(rows[j]["extracted_text"]) else j
                dup = j if master == i else i
                master_indices.setdefault(master, []).append((dup, score))
    # annotate
    for idx, row in enumerate(rows):
        row["duplicates"] = ""
        row["duplicates_percentage"] = ""
        row["master_one"] = ""

    for master, dup_list in master_indices.items():
        rows[master]["master_one"] = "MASTER"
        for dup_idx, score in dup_list:
            rows[dup_idx]["duplicates"] = f"Duplicate of {rows[master]['file_name']}"
            rows[dup_idx]["duplicates_percentage"] = f"{score*100:.2f}%"
            rows[dup_idx]["master_one"] = rows[master]["file_name"]


def process_folder(folder: Path, output_excel: Path):
    files = [p for p in folder.glob("**/*") if p.is_file() and p.suffix.lower() in {".pdf", ".docx", ".doc", ".pptx", ".ppt", ".msg"}]
    rows = []
    for f in files:
        try:
            text = extract_text_from_file(f)
        except Exception as e:
            print(f"Error extracting {f}: {e}")
            continue
        summary, file_type, version, tags = summarize_text(text)
        rows.append({
            "file_name": f.name,
            "file_type": file_type or f.suffix.lstrip('.'),
            "extracted_text": text,
            "summary": summary,
            "version": version,
            "tags": tags,
        })
    detect_duplicates(rows)
    df = pd.DataFrame(rows)
    df.to_excel(output_excel, index=False)
    print(f"Saved results to {output_excel}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Process documents and summarize via Azure OpenAI")
    parser.add_argument("--folder", required=True, help="Folder containing documents")
    parser.add_argument("--output", default="output.xlsx", help="Output excel path")
    args = parser.parse_args()

    process_folder(Path(args.folder), Path(args.output))