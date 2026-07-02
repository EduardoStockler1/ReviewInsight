from pathlib import Path

#arquivo criado somente pra limpar o pdf caso haja alguma informação irrelevante ou repetida, o objetivo é deixar o arquivo mais limpo e organizado para que posteriormente seja possível criar uma inteligência artificial que analise os dados e gere gráficos e insights a partir das avaliações coletadas

try:
    from PyPDF2 import PdfReader
except ImportError:
    PdfReader = None

try:
    from pdfminer.high_level import extract_text
except ImportError:
    extract_text = None

ROOT_DIR = Path(__file__).resolve().parent.parent
INPUT_PATH = ROOT_DIR / "hospital-erastinho" / "avaliacoes" / "scrapping" / "avaliacoes.pdf"
OUTPUT_DIR = INPUT_PATH.parent / "avaliacoes-limpo.pdf"
OUTPUT_PATH = OUTPUT_DIR / "cleaned_avaliacoes.txt"


def normalize_line(line: str) -> str:
    return " ".join(line.strip().split())


def is_short_comment(line: str) -> bool:
    return len(line.split()) < 5


def extract_pdf_text(path: Path) -> str:
    if PdfReader is not None:
        try:
            reader = PdfReader(path)
            pages = [page.extract_text() or "" for page in reader.pages]
            return "\n".join(pages)
        except Exception:
            pass
    if extract_text is not None:
        return extract_text(str(path))
    raise RuntimeError("No PDF extractor available. Install PyPDF2 or pdfminer.six.")


def clean_text(raw_text: str) -> str:
    seen = set()
    cleaned_lines = []
    for line in raw_text.splitlines():
        normalized = normalize_line(line)
        if not normalized:
            continue
        if is_short_comment(normalized):
            continue
        key = normalized.lower()
        if key in seen:
            continue
        seen.add(key)
        cleaned_lines.append(normalized)
    return "\n".join(cleaned_lines)


def main() -> None:
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"PDF not found: {INPUT_PATH}")

    raw_text = extract_pdf_text(INPUT_PATH)
    cleaned = clean_text(raw_text)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(cleaned, encoding="utf-8")
    print(f"Cleaned data saved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
