#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Skavenizer — takes text/document and "skavenizes" it:
  • duplicates some words with a hyphen (yes-yes, no-no);
  • adds a synonym to some words via hyphen;
  • adds a suffix to ANIMATE words:
       rus.:  <root>о-тварь / <root>о-твари
       eng.:  <word (sg.)>-thing / <word (pl. → sg.)>-things
       (plural → things, singular → thing; suffix matches number)

Russian root extraction modes:
  • pymystem3 (if installed and starts) — lemmatize + strip ending:
        человек → человеко-тварь, кошки → кошко-твари
  • fallback (without pymystem3) — first 4 letters minus trailing vowels:
        человек → чело-тварь, кошки → кошко-твари

English plural modes:
  • inflect (if installed) — correctly handles irregular:
        mice → mouse-things, children → child-things
        invariants (series, news, sheep, …) treated as plural → -things
  • fallback — suffix heuristic (s/es/ies) + invariant-word list.

English safeguard filters:
  • WordNet synonym lookup uses exact lemma match (no morphy), skips
    proper nouns for lowercase inputs, skips multi-word and uppercase
    synonyms, skips drastic length changes.
  • Function-word stop-list prevents pronouns/articles/auxiliaries
    (it, the, to, any, …) from getting a synonym or a "thing" suffix.
  • Adjective/adverb suffixes (-ous, -ious, -ful, -less, …) are excluded
    from plural detection to avoid garbage stems ("laborious" → "laboriou").

Limitations:
  • PDF with Cyrillic parses poorly (backend limitation of pypdf/PyPDF2).
  • .doc via textract requires antiword + libmagic on Windows.
  • Empty input exits with code 0 — not treated as an error.

Usage examples:
    python skavenizer.py input.docx
    python skavenizer.py "Человек идёт домой"
    python skavenizer.py input.pdf -o out.txt --seed 42
    echo "some text" | python skavenizer.py -
"""

from __future__ import annotations

import argparse
import random
import re
import sys
from pathlib import Path

# ─── Constants ───
RU_BEAST_SG = "тварь"
RU_BEAST_PL = "твари"
EN_BEAST_SG = "thing"
EN_BEAST_PL = "things"

ROOT_MAX_LEN = 4
ROOT_MIN_LEN = 2

DEFAULT_WEIGHTS = (0.15, 0.15, 0.15, 0.55)
_DEFAULT_WEIGHTS_LIST = list(DEFAULT_WEIGHTS)

# English invariant nouns: plural form equals singular form.
# The fallback branch (no inflect) treats them as plural → suffix -things.
_EN_INVARIANT_PLURALS = frozenset({
    "series", "species", "news",
    "fish", "sheep", "deer", "moose",
    "aircraft", "means", "offspring",
    "salmon", "trout", "swine",
})

# Adjective/adverb suffixes: inflect sometimes falsely strips a trailing -s
# from these ("laborious" → "laboriou"), producing garbage stems.
# Suffixes that look like they could hide a plural -s but don't.
# Prevents inflect from stripping the wrong letter ("happiness" → "happines",
# "laborious" → "laboriou", "witness" → "witnes").
_EN_NON_PLURAL_SUFFIXES = (
    # adjectives / adverbs
    "ous", "ious", "eous", "uous",
    "ful", "less", "ive", "able", "ible", "ly",
    # nouns in the singular
    "ness", "ship", "hood", "ment", "tion", "sion",
    # other common endings that end in -s but aren't plural
    "ss", "us", "is", "os",
)

# English function words: pronouns, articles, prepositions, conjunctions,
# auxiliaries, particles. These are never content nouns; never give them
# a synonym or a "thing" suffix, only duplicate/keep.
_EN_STOPWORDS = frozenset({
    # articles
    "a", "an", "the",
    # pronouns
    "i", "me", "my", "mine", "myself",
    "you", "your", "yours", "yourself", "yourselves",
    "he", "him", "his", "himself",
    "she", "her", "hers", "herself",
    "it", "its", "itself",
    "we", "us", "our", "ours", "ourselves",
    "they", "them", "their", "theirs", "themselves",
    "who", "whom", "whose", "which", "what",
    "this", "that", "these", "those",
    "one", "ones", "someone", "anyone", "everyone", "nobody",
    "somebody", "anybody", "everybody",
    "something", "anything", "everything", "nothing",
    # prepositions / particles
    "in", "on", "at", "by", "for", "with", "about", "against",
    "between", "into", "through", "during", "before", "after",
    "above", "below", "to", "from", "up", "down", "out", "off",
    "over", "under", "again", "further", "then", "once",
    "here", "there", "when", "where", "why", "how",
    "all", "any", "both", "each", "few", "more", "most", "other",
    "some", "such", "no", "nor", "not", "only", "own", "same",
    "so", "than", "too", "very",
    # conjunctions
    "and", "but", "or", "yet",
    "if", "else", "as", "because", "while", "until", "unless",
    "whether", "though", "although", "since",
    # auxiliaries
    "am", "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "having",
    "do", "does", "did", "doing", "done",
    "will", "would", "shall", "should",
    "may", "might", "must", "can", "could",
    # prepositions / particles
    "except", "despite", "upon", "within", "without",
    "along", "across", "beyond", "toward", "towards",
    "among", "amongst", "around", "behind", "beneath",
    "beside", "besides", "onto", "opposite", "outside",
    "past", "regarding", "throughout", "underneath",
    "unlike", "versus", "via", "considering",
})

# Lemma cache (pymystem3)
_lemma_cache: dict[str, str] = {}

# One-time warnings about missing synonym engines (per language)
_warned_no_syn = {"ru": False, "en": False}


# ─── File-reading libraries ───
try:
    from docx import Document
except ImportError:
    Document = None

try:
    from pypdf import PdfReader as _PdfReader
    _PDF_LIB = "pypdf"
except ImportError:
    try:
        from PyPDF2 import PdfReader as _PdfReader
        _PDF_LIB = "PyPDF2"
    except ImportError:
        _PdfReader = None
        _PDF_LIB = None

try:
    import textract
except ImportError:
    textract = None

# ─── Synonyms ───
try:
    from ruwordnet import RuWordNet
    _ru_wordnet = RuWordNet()
except Exception:
    _ru_wordnet = None

try:
    from nltk.corpus import wordnet as wn
except Exception:
    wn = None

# ─── Morphology: pymorphy3 ───
try:
    import pymorphy3
    _morph = pymorphy3.MorphAnalyzer()
except Exception:
    _morph = None

# ─── Morphology: pymystem3 (lazy init) ───
try:
    from pymystem3 import Mystem as _MystemClass
    _mystem_available = True
except Exception:
    _MystemClass = None
    _mystem_available = False

_mystem_instance = None


def _get_mystem():
    """Lazy Mystem() initialization — only on the first Russian request."""
    global _mystem_instance, _mystem_available
    if not _mystem_available:
        return None
    if _mystem_instance is None:
        try:
            _mystem_instance = _MystemClass()
        except Exception:
            _mystem_available = False
            return None
    return _mystem_instance

# ─── Morphology: inflect ───
try:
    import inflect as _inflect_mod
    _inflect = _inflect_mod.engine()
except Exception:
    _inflect = None


# ─────────────────────────────────────────────
# 1. FILE READING
# ─────────────────────────────────────────────

def _decode_bytes(data: bytes) -> str:
    """
    Try to decode bytes in order: utf-8 → cp1251 → latin-1.
    latin-1 never raises UnicodeDecodeError, so it's effectively the
    last resort — but it produces garbage for multibyte encodings,
    hence it goes last.
    """
    for enc in ("utf-8", "cp1251"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1")


def _read_text_file(file_path: str) -> str:
    """Read a text file with encoding fallbacks."""
    return _decode_bytes(Path(file_path).read_bytes())


def extract_text(file_path: str) -> str:
    """Extract text from a file by extension (pypdf / PyPDF2 / python-docx / textract)."""
    ext = Path(file_path).suffix.lower()

    if ext == ".txt":
        return _read_text_file(file_path)

    if ext == ".docx":
        if Document is None:
            raise ImportError("Install python-docx: pip install python-docx")
        doc = Document(file_path)
        return "\n".join(p.text for p in doc.paragraphs)

    if ext == ".doc":
        if textract is None:
            raise ImportError("Install textract: pip install textract")
        return _decode_bytes(textract.process(file_path))

    if ext == ".pdf":
        if _PdfReader is None:
            raise ImportError("Install pypdf (or PyPDF2): pip install pypdf")
        text = []
        with open(file_path, "rb") as f:
            reader = _PdfReader(f)
            for page in reader.pages:
                text.append(page.extract_text() or "")
        return "\n".join(text)

    raise ValueError(f"Unsupported extension: {ext}")


# ─────────────────────────────────────────────
# 2. WORDNET HELPERS
# ─────────────────────────────────────────────

def _wordnet_exact_synsets(word: str, pos=None):
    """
    Return WordNet synsets where `word` appears as an EXACT lemma
    (case-insensitive), suppressing NLTK's built-in morphy.
    Also skips proper-noun lemmas (Born, Washington) unless the input
    word itself was capitalized.

    Prevents "was" → "wa" (Washington), "has" → "ha", "who" → "WHO",
    "born" → "Born (Max Born)", etc.
    """
    if wn is None:
        return []
    w = word.lower()
    was_capitalized = bool(word) and word[:1].isupper()
    try:
        raw = wn.synsets(w, pos=pos) if pos else wn.synsets(w)
    except Exception:
        return []
    out = []
    for s in raw:
        for lemma in s.lemmas(lang="eng"):
            name = lemma.name()
            if name.lower() != w:
                continue
            if name[:1].isupper() and not was_capitalized:
                continue  # proper noun, but input was lowercase
            out.append(s)
            break
    return out

def _is_person_or_animal_synset(s) -> bool:
    """
    True if the synset is (or inherits from) noun.person / noun.animal.
    Uses a fast lexname check first, then a hypernym walk for robustness.
    """
    try:
        if s.lexname() in ("noun.animal", "noun.person"):
            return True
    except Exception:
        return False
    try:
        for path in s.hypernym_paths():
            for h in path:
                if h.name() in ("person.n.01", "animal.n.01"):
                    return True
    except Exception:
        pass
    return False

# ─────────────────────────────────────────────
# 3. ANIMACY AND NUMBER
# ─────────────────────────────────────────────

def is_animate_ru(word: str) -> bool:
    if _morph is None:
        return True
    try:
        return "anim" in _morph.parse(word.lower())[0].tag
    except Exception:
        return True


def is_animate_en(word: str) -> bool:
    """
    Is the English word animate?
      • Word in WordNet but not a noun       → False (verbs, adjectives, …).
      • Word not in WordNet at all           → True  (unknown, lenient).
      • Otherwise: judged by the FIRST (most common) noun sense and its
        hypernym path to person/animal. Avoids false positives like
        "love" (love.n.01 = feeling, even though love.n.03 = beloved person).
    """
    if wn is None:
        return True
    try:
        exact = _wordnet_exact_synsets(word)
        if not exact:
            return True
        noun_synsets = _wordnet_exact_synsets(word, pos=wn.NOUN)
        if not noun_synsets:
            return False
        return _is_person_or_animal_synset(noun_synsets[0])
    except Exception:
        return True

def is_plural_ru(word: str) -> bool:
    if _morph is not None:
        try:
            return "plur" in _morph.parse(word.lower())[0].tag
        except Exception:
            pass
    return word.lower().endswith(("ы", "и"))


def _is_plural_en_fallback(word: str) -> bool:
    """Suffix heuristic (no inflect). Invariants → treated as plural."""
    w = word.lower()
    if w in _EN_INVARIANT_PLURALS:
        return True
    if w.endswith(("ss", "us", "is", "os")):
        return False
    return w.endswith("s")


def is_plural_en(word: str) -> bool:
    w = word.lower()
    if w.endswith(_EN_NON_PLURAL_SUFFIXES):
        return False
    if _inflect is not None:
        try:
            return _inflect.singular_noun(word) is not False
        except Exception:
            pass
    return _is_plural_en_fallback(word)


# ─────────────────────────────────────────────
# 4. SYNONYM LOOKUP
# ─────────────────────────────────────────────

def get_synonym(word: str, lang: str = "ru") -> str | None:
    word_lower = word.lower()

    if lang == "ru":
        if _ru_wordnet is None:
            if not _warned_no_syn["ru"]:
                print("[Skavenizer] ruwordnet is unavailable — Russian synonyms "
                      "will not be picked. Try: ruwordnet download",
                      file=sys.stderr)
                _warned_no_syn["ru"] = True
            return None
        try:
            for synset in _ru_wordnet.get_synsets(word_lower):
                for sense in synset.senses:
                    name = sense.name
                    if name.lower() == word_lower:
                        continue
                    if " " in name:
                        continue           # skip multi-word synonyms
                    if not name[:1].islower():
                        continue           # skip proper nouns / acronyms
                    if len(name) > len(word) * 2 + 2:
                        continue           # skip drastic length changes
                    return name
        except Exception as e:
            if not _warned_no_syn["ru"]:
                print(f"[Skavenizer] ruwordnet: {e}", file=sys.stderr)
                _warned_no_syn["ru"] = True

    if lang == "en":
        if wn is None:
            if not _warned_no_syn["en"]:
                print("[Skavenizer] wordnet is unavailable — English synonyms "
                      "will not be picked. Install: "
                      "python -c \"import nltk; nltk.download('wordnet')\"",
                      file=sys.stderr)
                _warned_no_syn["en"] = True
            return None
        try:
            # Use the ORIGINAL word (not lowercased) so that
            # _wordnet_exact_synsets can honour capitalization correctly.
            for synset in _wordnet_exact_synsets(word):
                for lemma in synset.lemmas(lang="eng"):
                    name = lemma.name().replace("_", " ")
                    if name.lower() == word_lower:
                        continue
                    if " " in name:
                        continue           # skip multi-word synonyms
                    if not name[:1].islower():
                        continue           # skip proper names
                    if len(name) > len(word) * 2 + 2:
                        continue           # skip drastic length changes
                    return name
        except Exception:
            pass

    return None


# ─────────────────────────────────────────────
# 5. RUSSIAN ROOT EXTRACTION
# ─────────────────────────────────────────────

# Endings sorted by length descending. Diminutive suffixes (-ек/-ок/-ик/-ец/-иц)
# are INTENTIONALLY EXCLUDED: they cause false positives (человек → челов),
# and diminutive roots are already handled by pymystem3 lemmatization.
_RU_ENDINGS = tuple(dict.fromkeys((
    # 4
    "иями",
    # 3
    "ями", "ами", "ией", "иях", "иям", "ого", "ему", "ому", "ыми", "ими",
    # 2
    "ию", "ия", "ие", "ые", "ых", "ей", "ой", "ий", "ый",
    "ая", "ое", "ее", "ам", "ах", "ов", "ев",
    # 1
    "у", "ю", "ы", "и", "а", "я", "о", "е", "ь",
)))

_RU_VOWELS = "аеёиоуыэюя"


def lemma_ru(word: str) -> str:
    mystem = _get_mystem()
    if mystem is None:
        return word
    w = word.lower()
    cached = _lemma_cache.get(w)
    if cached is not None:
        return cached
    try:
        lemmas = mystem.lemmatize(w)
        lemma = lemmas[0] if lemmas else w
        if not re.fullmatch(r"[а-яё]+", lemma, re.IGNORECASE):
            lemma = w
    except Exception:
        lemma = w
    _lemma_cache[w] = lemma
    return lemma


def russian_root_pymystem(word: str) -> str:
    lemma = lemma_ru(word)
    stem = lemma
    for end in _RU_ENDINGS:
        if stem.endswith(end) and len(stem) - len(end) >= ROOT_MIN_LEN:
            stem = stem[: -len(end)]
            break
    return stem


def russian_root_heuristic(word: str) -> str:
    n = min(ROOT_MAX_LEN, len(word))
    stem = word[:n]
    while len(stem) > ROOT_MIN_LEN and stem and stem[-1].lower() in _RU_VOWELS:
        stem = stem[:-1]
    return stem


def russian_root(word: str) -> str:
    if _get_mystem() is not None:
        return russian_root_pymystem(word)
    return russian_root_heuristic(word)


def apply_case(source: str, target: str) -> str:
    """
    Apply source's case pattern to target:
      • ALL UPPER  → target UPPER
      • Capitalized → target Capitalized
      • lowercase   → target lowercase
    """
    if not source:
        return target
    if source.isupper():
        return target.upper()
    if source[:1].isupper():
        return target[:1].upper() + target[1:].lower()
    return target.lower()


def make_beast_ru(word: str) -> str:
    plural = is_plural_ru(word)
    root = russian_root(word)
    if not root:
        return word
    join = "-" if root[-1].lower() in _RU_VOWELS else "о-"
    suffix = RU_BEAST_PL if plural else RU_BEAST_SG
    return apply_case(word, f"{root}{join}{suffix}")


# ─────────────────────────────────────────────
# 6. ENGLISH SINGULAR / PLURAL
# ─────────────────────────────────────────────

def strip_plural_en_fallback(word: str) -> str:
    lower = word.lower()
    if lower in _EN_INVARIANT_PLURALS:
        return word
    if lower.endswith("ies") and len(word) > 4:
        return word[:-3] + "y"
    if lower.endswith(("ses", "xes", "zes", "ches", "shes")) and len(word) > 3:
        return word[:-2]
    if lower.endswith("s") and len(word) > 1 and not lower.endswith(
        ("ss", "us", "is", "os")
    ):
        return word[:-1]
    return word


def singular_en(word: str) -> str:
    """Return the singular form of an English word. Preserves case."""
    if _inflect is not None:
        try:
            s = _inflect.singular_noun(word)
            if s:
                return apply_case(word, s)
        except Exception:
            pass
    return strip_plural_en_fallback(word)


def make_beast_en(word: str, plural: bool, sg: str) -> str:
    """
    plural → <sg>-things, singular → <word>-thing.
    `plural` and `sg` are computed by the caller (no repeated calls).
    """
    stem = sg if plural else word
    suffix = EN_BEAST_PL if plural else EN_BEAST_SG
    return f"{stem}-{suffix}"


# ─────────────────────────────────────────────
# 7. WORD-LEVEL SKAVENIZATION
# ─────────────────────────────────────────────

def skavenize_word(word: str, weights: list[float] | None = None) -> str:
    if not word or not word[:1].isalpha():
        return word

    w = weights if weights is not None else _DEFAULT_WEIGHTS_LIST
    is_cyrillic = bool(re.search(r"[а-яА-ЯёЁ]", word))
    actual_lang = "ru" if is_cyrillic else "en"
    is_stopword = (not is_cyrillic) and word.lower() in _EN_STOPWORDS

    choice = random.choices(
        ["duplicate", "synonym", "tvari", "keep"],
        weights=w, k=1,
    )[0]

    if choice == "duplicate":
        return f"{word}-{word}"

    if choice == "synonym":
        if is_stopword:
            return word
        syn = get_synonym(word, actual_lang)
        if syn:
            return f"{word}-{apply_case(word, syn)}"
        return word

    if choice == "tvari":
        if is_stopword:
            return word
        if is_cyrillic:
            return make_beast_ru(word) if is_animate_ru(word) else word
        plural = is_plural_en(word)
        sg = singular_en(word) if plural else word
        if not is_animate_en(sg):
            return word
        return make_beast_en(word, plural, sg)

    return word


# ─────────────────────────────────────────────
# 8. TEXT-LEVEL SKAVENIZATION
# ─────────────────────────────────────────────

def skavenize_text(text: str, weights: list[float] | None = None) -> str:
    tokens = re.split(r"(\s+|[^\w\s]+)", text)
    result = []
    for token in tokens:
        if re.fullmatch(r"\w+", token, re.UNICODE):
            result.append(skavenize_word(token, weights))
        else:
            result.append(token)
    return "".join(result)


# ─────────────────────────────────────────────
# 9. CLI AND MAIN
# ─────────────────────────────────────────────

def _default_output_for(source: str, is_file: bool) -> str:
    """
    If source was a real file — '<stem>_skavenized.txt'.
    Otherwise (raw text, stdin) — 'skavenized_output.txt'.
    """
    if not is_file or source == "-":
        return "skavenized_output.txt"
    return f"{Path(source).stem}_skavenized.txt"


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Skavenizer — duplicates/replaces/tvarizes words.",
    )
    p.add_argument("source",
                   help="File path (.txt/.docx/.doc/.pdf), raw text, or '-' for stdin")
    p.add_argument("-o", "--output", default=None,
                   help="Where to save the result "
                        "(default: <input_stem>_skavenized.txt for a file, "
                        "otherwise skavenized_output.txt)")
    p.add_argument("--seed", type=int, default=None,
                   help="random seed — for reproducibility")
    p.add_argument("--weights", type=float, nargs=4,
                   metavar=("DUP", "SYN", "TVA", "KEEP"),
                   default=list(DEFAULT_WEIGHTS),
                   help="Weights for duplicate synonym tvari keep "
                        f"(default: {' '.join(map(str, DEFAULT_WEIGHTS))})")
    p.add_argument("-q", "--quiet", action="store_true",
                   help="File only: no result output and no diagnostics "
                        "(error messages still go to stderr)")
    p.add_argument("--preview", type=int, default=0, metavar="N",
                   help="Print only the first N characters (0 = entire text)")

    args = p.parse_args()

    if sum(args.weights) <= 0 or any(w < 0 for w in args.weights):
        p.error("Weights must be non-negative and sum to a value > 0")

    return args


def _print_diagnostics(quiet: bool) -> None:
    if quiet:
        return

    # pymystem3: force lazy init so we don't lie about its status
    if _get_mystem() is None:
        print("[Skavenizer] pymystem3 is unavailable or failed to start — "
              "Russian root via fallback heuristic (first 4 letters). "
              "Install: pip install pymystem3",
              file=sys.stderr)
    else:
        print("[Skavenizer] pymystem3 is active — Russian roots via lemma.",
              file=sys.stderr)

    if _inflect is None:
        print("[Skavenizer] inflect not found — irregular plurals "
              "(mice, children) will not be recognized. Install: pip install inflect",
              file=sys.stderr)
    else:
        print("[Skavenizer] inflect is active — English plurals handled correctly.",
              file=sys.stderr)

    if _PDF_LIB is not None:
        print(f"[Skavenizer] PDF backend: {_PDF_LIB}.", file=sys.stderr)


def main() -> None:
    args = parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    _print_diagnostics(args.quiet)

    # Source: stdin, file, or raw text. Remember whether it was a file,
    # so _default_output_for doesn't confuse "Hello. World." (raw text)
    # with a file path.
    is_file = False
    if args.source == "-":
        if not args.quiet:
            print("[Skavenizer] Reading stdin.", file=sys.stderr)
        raw_text = sys.stdin.read()
    else:
        path = Path(args.source)
        is_file = path.exists() and path.is_file()
        if is_file:
            if not args.quiet:
                print(f"[Skavenizer] Reading file: {path}", file=sys.stderr)
            raw_text = extract_text(str(path))
        else:
            if not args.quiet:
                print("[Skavenizer] Processing text from argument.",
                      file=sys.stderr)
            raw_text = args.source

    if not raw_text.strip():
        print("Text is empty.", file=sys.stderr)
        sys.exit(0)  # empty input is not an error

    skavenized = skavenize_text(raw_text, weights=list(args.weights))

    # Where to save
    out_arg = args.output or _default_output_for(args.source, is_file)
    out_path = Path(out_arg)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(skavenized, encoding="utf-8")
    if not args.quiet:
        print(f"[Skavenizer] Result saved to: {out_path.resolve()}",
              file=sys.stderr)

    if not args.quiet:
        to_print = skavenized
        if args.preview > 0:
            to_print = to_print[: args.preview]
            if len(skavenized) > args.preview:
                to_print += f"\n…[{len(skavenized) - args.preview} more characters]"
        print("\n" + "=" * 60)
        print(to_print)
        print("=" * 60 + "\n")


if __name__ == "__main__":
    main()