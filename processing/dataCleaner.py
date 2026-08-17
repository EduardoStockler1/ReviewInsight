"""
Limpeza completa de avaliacoes (reviews) extraidas do Google/etc.

Como usar
---------
1. Ajuste `file_path` em main() para o caminho do seu avaliacoes.xlsx.
2. Ajuste o dicionario CONFIG abaixo para ligar/desligar as limpezas opcionais.
3. Rode: python limpar_avaliacoes.py

O que este script cobre
------------------------
Quase item por item da sua lista. Os itens marcados como "(opcional)" na sua
lista viraram flags no CONFIG (todas comecam desligadas, exceto as que sao
"seguras" por padrao, tipo emoji/emoticon/GIF/mojibake).

Itens que NAO da para automatizar com regex de forma confiavel (correcao
ortografica, expansao de girias/abreviacoes, deteccao de "texto de sistema"
generico, remocao de "rodapes/cabecalhos" arbitrarios) ficam como GANCHOS:
funcoes/dicionarios vazios prontos para voce preencher com as regras do seu
dominio, em vez de eu inventar um comportamento que pode apagar dado bom.
"""

import pandas as pd
import numpy as np
import re
import os
import unicodedata

# =====================================================================
# CONFIG — ligue/desligue o que quiser aqui
# =====================================================================
CONFIG = {
    # Colunas
    # Nome (case-insensitive) da coluna com o texto do comentario/review.
    # Se None, o script tenta adivinhar procurando por nomes comuns.
    "comment_column": None,

    # Limpezas "sempre ligadas" (fazem parte do minimo esperado)
    # -> HTML/XML, URLs, e-mails, telefone, CPF/CNPJ/CEP, invisiveis,
    #    controle, mojibake/unicode invalido, espacos, pontuacao repetida,
    #    linhas/colunas vazias, duplicatas -> essas sao sempre aplicadas.

    # Limpezas opcionais (default = False, exceto indicado)
    "remove_accents": False,              # Acentos (opcional)
    "lowercase": False,                   # Normalizacao para minusculas (opcional)
    "remove_emojis": True,                # Emojis (opcional) - default ligado
    "remove_emoticons": True,             # Emoticons (opcional) - default ligado
    "remove_mentions": True,              # @usuario, se irrelevantes
    "remove_hashtags": True,              # Hashtags, se irrelevantes
    "remove_stopwords": False,            # Stopwords (opcional)
    "reduce_letter_repetition": True,     # boooom -> boom (opcional) - default ligado
    "reduce_word_repetition": True,       # bom bom bom -> bom (opcional) - default ligado
    "remove_markdown": True,              # **, __, `, # quando irrelevantes
    "expand_abbreviations": False,        # usa ABBREVIATIONS_MAP abaixo
    "normalize_slang": False,             # usa SLANG_MAP abaixo
    "spell_correct": False,               # gancho — nao implementado (ver nota)

    # Filtros de linha
    "min_comment_length": 3,              # 0 = desliga o filtro de tamanho minimo
    "drop_comment_only_numbers_symbols": True,
    "drop_comment_only_links": True,
    "drop_comment_only_emojis": True,     # so tem efeito antes de remover emojis
    "drop_comment_only_whitespace_or_invisible": True,
    "drop_duplicate_rows": True,
    "drop_duplicate_comments": True,      # duplicatas exatas de comentario
    "drop_fully_empty_columns": True,
    "irrelevant_columns": [],             # ex: ["id_interno", "url_foto"]

    # Placeholders que devem virar NaN (N/A, -, Sem resposta, etc.)
    "placeholder_values": [
        "n/a", "na", "-", "--", "s/n", "sem resposta", "nao informado",
        "não informado", "null", "none", "undefined", ".", "..", "...",
    ],
}

# Ganchos para voce preencher (opcional)
ABBREVIATIONS_MAP = {
    # "vc": "voce", "pq": "porque", "tb": "tambem",
}
SLANG_MAP = {
    # "mds": "meu deus", "top": "otimo",
}

# =====================================================================
# Utilitarios de baixo nivel
# =====================================================================

INVISIBLE_CHARS_RE = re.compile(
    "["
    "\u200b\u200c\u200d\u200e\u200f"   # zero-width space/joiners/marks
    "\u2060\ufeff"                      # word joiner, BOM
    "\u00ad"                            # soft hyphen
    "]"
)

CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

HTML_TAG_RE = re.compile(r"<[^>]+>")
XML_TAG_RE = re.compile(r"</?[a-zA-Z][\w:-]*(?:\s+[^<>]*)?/?>")
SCRIPT_STYLE_RE = re.compile(r"<(script|style)[^>]*>.*?</\1>", re.DOTALL | re.IGNORECASE)

URL_RE = re.compile(
    r"(?:https?://|www\.)[^\s<>\"']+", re.IGNORECASE
)
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

# Telefones BR (fixo/celular, com/sem DDD, com/sem +55)
PHONE_RE = re.compile(
    r"(?:\+?55\s?)?(?:\(?\d{2}\)?\s?)?9?\d{4}[-.\s]?\d{4}"
)

CPF_RE = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}-?\d{2}\b")
CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
CEP_RE = re.compile(r"\b\d{5}-?\d{3}\b")
# RG nao tem formato nacional unico; cobre os padroes mais comuns (SP/MG/RJ etc.)
RG_RE = re.compile(r"\b\d{1,2}\.?\d{3}\.?\d{3}-?[0-9xX]\b")

MENTION_RE = re.compile(r"(?<!\w)@\w+")
HASHTAG_RE = re.compile(r"(?<!\w)#\w+")

# Cobre emoticons "texto" tipo :) :( :D XD ^^ e tambem emojis unicode
EMOTICON_RE = re.compile(
    r"(?::|;|=|X|x)[\-o\*']?(?:\)|\(|D|P|p|O|o|\/|\\|3|\||S|s)"
    r"|<3|\)-?:|\(-?:|\^_?\^|xD|XD"
)

EMOJI_RE = re.compile(
    "["
    "\U0001F300-\U0001FAFF"
    "\U00002600-\U000027BF"
    "\U0001F1E6-\U0001F1FF"
    "\U00002190-\U000021FF"
    "\U00002B00-\U00002BFF"
    "\U0001F900-\U0001F9FF"
    "\U0000FE0F"
    "]+"
)

GIF_IMG_PLACEHOLDER_RE = re.compile(
    r"\[?\s*(?:gif|imagem|image|photo|foto|video|attachment|anexo)\s*\]?",
    re.IGNORECASE,
)

MARKDOWN_RE = re.compile(r"(\*\*|__|`{1,3}|~~|^#{1,6}\s?)", re.MULTILINE)

REPEATED_PUNCT_RE = re.compile(r"([!?.,;:])\1{1,}")
MULTI_QUESTION_EXCLAM_RE = re.compile(r"([!?])(?:\s*[!?])+")

BACKSLASH_RE = re.compile(r"\\+(?!n|t|r)")
LITERAL_ESCAPE_RE = re.compile(r"\\[ntr]")

ISOLATED_SYMBOL_RE = re.compile(
    r"(?<!\w)[§¤©®™°¶†‡•·…»«¬~^`´¨]+(?!\w)"
)

LETTER_REPEAT_RE = re.compile(r"(\p{L})\1{2,}", re.UNICODE) if False else re.compile(
    r"([a-zA-ZáéíóúâêôãõàçÁÉÍÓÚÂÊÔÃÕÀÇ])\1{2,}"
)

# duas ou mais ocorrencias seguidas da MESMA palavra (>=2 letras)
WORD_REPEAT_RE = re.compile(r"\b(\w{2,})(?:\s+\1\b)+", re.IGNORECASE)

HTML_ENTITY_MAP = {
    "&nbsp;": " ", "&amp;": "&", "&quot;": '"', "&lt;": "<", "&gt;": ">",
    "&apos;": "'", "&#39;": "'", "&#160;": " ",
}

STOPWORDS_PT = {
    "a", "o", "as", "os", "de", "do", "da", "dos", "das", "em", "um", "uma",
    "uns", "umas", "e", "é", "ao", "aos", "à", "às", "para", "por", "com",
    "que", "se", "na", "no", "nas", "nos", "sua", "seu", "suas", "seus",
    "esse", "essa", "isso", "este", "esta", "isto", "mas", "ou", "como",
    "foi", "ser", "ter", "muito", "mais", "ja", "já", "tambem", "também",
}


def remove_accents(text):
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


def fix_mojibake(text):
    """
    Corrige o caso mais comum de mojibake (texto UTF-8 lido como Latin-1,
    tipo 'Ã©' no lugar de 'é'), sem depender de bibliotecas externas.
    Se a tentativa de round-trip falhar ou nao mudar nada plausivel,
    mantem o texto original.
    """
    if any(ch in text for ch in ("Ã", "Â", "â€", "�")):
        try:
            fixed = text.encode("latin-1").decode("utf-8")
            return fixed
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text
    return text


def normalize_invalid_unicode(text):
    # NFKC junta formas de compatibilidade (ex: largura total) em formas normais
    # e remove caracteres nao atribuidos/():
    text = unicodedata.normalize("NFKC", text)
    return "".join(c for c in text if unicodedata.category(c) != "Cn")


def is_only_symbols_or_numbers(text):
    stripped = text.strip()
    if not stripped:
        return True
    # nada de letras (unicode) no texto
    return not re.search(r"[^\W\d_]", stripped, re.UNICODE)


def is_only_links(text):
    stripped = URL_RE.sub("", text).strip()
    return stripped == "" and bool(URL_RE.search(text))


def is_only_whitespace_or_invisible(text):
    stripped = INVISIBLE_CHARS_RE.sub("", text)
    stripped = CONTROL_CHARS_RE.sub("", stripped)
    return stripped.strip() == ""


def is_only_emojis(text):
    stripped = EMOJI_RE.sub("", text).strip()
    return stripped == "" and bool(EMOJI_RE.search(text))


# =====================================================================
# Pipeline de limpeza de texto (aplicado celula a celula)
# =====================================================================

def clean_text(value, cfg):
    if pd.isna(value):
        return np.nan

    text = str(value)

    # 1) mojibake / unicode invalido primeiro, antes de qualquer regex
    text = fix_mojibake(text)
    text = normalize_invalid_unicode(text)

    # 2) invisiveis e caracteres de controle
    text = INVISIBLE_CHARS_RE.sub("", text)
    text = CONTROL_CHARS_RE.sub("", text)

    # 3) entidades HTML -> caractere real, depois remove tags HTML/XML e scripts/CSS
    for entity, char in HTML_ENTITY_MAP.items():
        text = text.replace(entity, char)
    text = re.sub(r"&#\d+;", " ", text)          # entidades numericas restantes
    text = SCRIPT_STYLE_RE.sub(" ", text)          # <script>...</script> / <style>...</style>
    text = HTML_TAG_RE.sub(" ", text)              # <p>, <br>, etc.
    text = XML_TAG_RE.sub(" ", text)               # tags XML residuais

    if cfg["remove_markdown"]:
        text = MARKDOWN_RE.sub(" ", text)

    # 4) sequencias de escape literais "\n" "\t" "\r" (texto, nao caractere real)
    text = LITERAL_ESCAPE_RE.sub(" ", text)
    # quebras de linha/tab reais -> espaco
    text = re.sub(r"[\r\n\t]+", " ", text)

    # 5) placeholders de imagem/gif/anexo
    text = GIF_IMG_PLACEHOLDER_RE.sub(" ", text)

    # 6) URLs, e-mails, documentos, telefones (nessa ordem: CPF/CNPJ/CEP antes
    #    de telefone, porque tem digitos parecidos e senao o regex de telefone
    #    "come" parte de um CPF, por exemplo)
    text = URL_RE.sub(" ", text)
    text = EMAIL_RE.sub(" ", text)
    text = CNPJ_RE.sub(" ", text)
    text = CPF_RE.sub(" ", text)
    text = CEP_RE.sub(" ", text)
    text = RG_RE.sub(" ", text)
    text = PHONE_RE.sub(" ", text)

    if cfg["remove_mentions"]:
        text = MENTION_RE.sub(" ", text)
    if cfg["remove_hashtags"]:
        text = HASHTAG_RE.sub(" ", text)

    if cfg["remove_emoticons"]:
        text = EMOTICON_RE.sub(" ", text)
    if cfg["remove_emojis"]:
        text = EMOJI_RE.sub(" ", text)

    # 7) pontuacao/aspas/barras
    text = REPEATED_PUNCT_RE.sub(r"\1", text)          # !!!! -> !   ...... -> .
    text = MULTI_QUESTION_EXCLAM_RE.sub(lambda m: m.group(1), text)
    text = re.sub(r'"{2,}', '"', text)                  # aspas duplicadas
    text = re.sub(r"'{2,}", "'", text)
    text = BACKSLASH_RE.sub(" ", text)                   # barras invertidas soltas
    text = ISOLATED_SYMBOL_RE.sub(" ", text)             # §, ¤, ©, ® isolados

    if cfg["reduce_letter_repetition"]:
        text = LETTER_REPEAT_RE.sub(r"\1\1", text)       # boooom -> boom

    if cfg["reduce_word_repetition"]:
        text = WORD_REPEAT_RE.sub(r"\1", text)           # bom bom bom -> bom

    # remove palavras duplicadas consecutivas (caso geral, curtas ou longas)
    text = re.sub(r"\b(\w+)\s+\1\b", r"\1", text, flags=re.IGNORECASE)

    if cfg["expand_abbreviations"] and ABBREVIATIONS_MAP:
        for abbr, full in ABBREVIATIONS_MAP.items():
            text = re.sub(rf"\b{re.escape(abbr)}\b", full, text, flags=re.IGNORECASE)

    if cfg["normalize_slang"] and SLANG_MAP:
        for slang, norm in SLANG_MAP.items():
            text = re.sub(rf"\b{re.escape(slang)}\b", norm, text, flags=re.IGNORECASE)

    # 8) espacos: multiplos -> um, trim
    text = re.sub(r"\s+", " ", text).strip()

    # 9) stopwords (opcional) - por ultimo, pra nao atrapalhar os regex acima
    if cfg["remove_stopwords"] and text:
        words = [w for w in text.split(" ") if w.lower() not in STOPWORDS_PT]
        text = " ".join(words)

    # 10) opcionais finais
    if cfg["remove_accents"]:
        text = remove_accents(text)
    if cfg["lowercase"]:
        text = text.lower()

    # texto em branco / so simbolos -> NaN
    if is_only_symbols_or_numbers(text) or text == "":
        return np.nan

    return text


# =====================================================================
# Helpers de coluna/linha
# =====================================================================

def guess_comment_column(df):
    candidates = ["comentario", "comentário", "review", "avaliacao",
                  "avaliação", "texto", "descricao", "descrição", "comment"]
    for col in df.columns:
        if col.strip().lower() in candidates:
            return col
    return None


def extract_rating(value):
    if pd.isna(value):
        return value
    match = re.search(r"\d+(?:[.,]\d+)?", str(value))
    if not match:
        return None
    return match.group().replace(",", ".")


def get_available_directory(base_path, base_name):
    dir_path = os.path.join(base_path, base_name)
    if not os.path.exists(dir_path):
        return dir_path
    counter = 1
    while True:
        new_dir_path = os.path.join(base_path, f"{base_name}{counter}")
        if not os.path.exists(new_dir_path):
            return new_dir_path
        counter += 1


# =====================================================================
# Pipeline principal
# =====================================================================

def clean_dataframe(df, cfg):
    df = df.copy()

    # colunas irrelevantes definidas manualmente
    if cfg["irrelevant_columns"]:
        df = df.drop(columns=[c for c in cfg["irrelevant_columns"] if c in df.columns])

    # coluna de nota (numero)
    rating_column = next((c for c in df.columns if c.strip().lower() == "nota"), None)
    if rating_column is not None:
        df[rating_column] = df[rating_column].apply(extract_rating)
        df[rating_column] = pd.to_numeric(df[rating_column], errors="coerce")

    # coluna de comentario
    comment_column = cfg["comment_column"] or guess_comment_column(df)

    # limpeza de texto em todas as colunas de texto
    # (cobre tanto dtype 'object' quanto o dtype 'string' do pandas >= 2,
    # que passou a ser o default em vez de 'object' em versoes recentes)
    text_columns = [
        c for c in df.columns
        if df[c].dtype == object or pd.api.types.is_string_dtype(df[c])
    ]
    for col in text_columns:
        df[col] = df[col].apply(lambda v: clean_text(v, cfg))

    # placeholders -> NaN (depois da limpeza de texto, comparando lowercase)
    placeholders_lower = {p.lower() for p in cfg["placeholder_values"]}
    for col in text_columns:
        df[col] = df[col].apply(
            lambda v: np.nan if (isinstance(v, str) and v.strip().lower() in placeholders_lower) else v
        )

    # colunas totalmente vazias
    if cfg["drop_fully_empty_columns"]:
        df = df.dropna(axis=1, how="all")

    # filtros baseados na coluna de comentario, se ela existir
    if comment_column is not None and comment_column in df.columns:
        if cfg["min_comment_length"] > 0:
            df = df[
                df[comment_column].isna()
                | (df[comment_column].astype(str).str.len() >= cfg["min_comment_length"])
            ]

        if cfg["drop_duplicate_comments"]:
            df = df.drop_duplicates(subset=[comment_column], keep="first")

    # linhas totalmente vazias (todas as colunas NaN)
    df = df.dropna(axis=0, how="all")

    # linhas duplicadas (considerando todas as colunas)
    if cfg["drop_duplicate_rows"]:
        df = df.drop_duplicates(keep="first")

    df = df.reset_index(drop=True)
    return df, comment_column


def main():
    file_path = r"C:\Users\stock\faculdade\projeto_iniciacao_cientifica\Review-Insight\ambrosia\avaliacoes.xlsx"
    df = pd.read_excel(file_path)

    print(f"Linhas antes da limpeza: {len(df)}")
    df_clean, comment_column = clean_dataframe(df, CONFIG)
    print(f"Linhas depois da limpeza: {len(df_clean)}")
    if comment_column:
        print(f"Coluna de comentario identificada: '{comment_column}'")
    else:
        print("Aviso: nao identifiquei automaticamente a coluna de comentario. "
              "Defina CONFIG['comment_column'] manualmente se quiser os filtros "
              "de comentario (tamanho minimo, duplicatas de comentario, etc.).")

    base_path = os.path.dirname(file_path)
    output_dir = get_available_directory(base_path, "cleaned_data")
    os.makedirs(output_dir, exist_ok=True)
    output_file_path = os.path.join(output_dir, "reviews_cleaned.xlsx")

    df_clean.to_excel(output_file_path, index=False)
    print("Limpeza concluida!")
    print(f"Arquivo limpo salvo em: {output_file_path}")


if __name__ == "__main__":
    main()