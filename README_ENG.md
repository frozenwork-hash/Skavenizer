# Skavenizer

A Python text-transformation utility that rewrites input documents by duplicating
words, appending synonyms, and attaching a language-aware "beast" suffix to
animate nouns. Designed for playful parody and linguistic experiments in both
Russian and English.

---

## Table of Contents

1. [Overview](#overview)
2. [How It Works](#how-it-works)
3. [Features](#features)
4. [Installation](#installation)
5. [Usage](#usage)
6. [Transformation Rules](#transformation-rules)
7. [Optional Engines](#optional-engines)
8. [Function Reference](#function-reference)
9. [Limitations](#limitations)
10. [Credits](#credits)

---

## Overview

Skavenizer reads text from a file (`.txt`, `.docx`, `.doc`, `.pdf`), from a
raw command-line argument, or from standard input, and produces a transformed
version of that text. The transformation applies one of four randomly weighted
operations to each word:

- **duplicate** — the word is repeated with a hyphen (`yes-yes`, `да-да`);
- **synonym** — a synonym is appended with a hyphen (`good-fine`);
- **tvari** — the word is reduced to its root and a language-specific
  "beast" suffix is appended (`чело-тварь`, `dog-thing`);
- **keep** — the word is left unchanged.

The "tvari" transformation is applied only to **animate** nouns, as determined
by morphological analysis (Russian) or WordNet lexicographer classes
(English). Number agreement is preserved: singular inputs receive a singular
suffix, plural inputs receive a plural suffix.

---

## How It Works

The pipeline consists of six stages:

1. **Input acquisition.** The source may be a filesystem path, a raw string,
   or `-` (standard input). File type is inferred from the extension.
2. **Tokenization.** The text is split into tokens using a regular expression
   that preserves whitespace and punctuation as separate tokens.
3. **Per-word transformation.** Each alphabetic token is passed through the
   weighted random selector.
4. **Language detection.** A word is classified as Russian if it contains at
   least one Cyrillic letter; otherwise it is treated as English.
5. **Animate check.** Russian animacy is determined via `pymorphy3` (grammeme
   `anim`); English animacy is determined via WordNet (`noun.animal`,
   `noun.person`).
6. **Reconstruction and output.** Tokens are rejoined and written to both a
   file and (unless suppressed) standard output.

---

## Features

- **Multi-format input.** `.txt`, `.docx`, `.doc`, `.pdf`, plus raw text and stdin.
- **Encoding resilience.** `.txt` files are decoded with automatic fallback:
  UTF-8 → CP1251 → Latin-1.
- **Bilingual support.** Russian and English, independently per word.
- **Number agreement.** `человек → человеко-тварь`, `люди → человеко-твари`;
  `dog → dog-thing`, `dogs → dog-things`.
- **Case preservation.** `Собака → Собако-тварь`, `СОБАКА → СОБАКО-ТВАРЬ`.
- **Optional morphological engines.** Pymystem3 (Russian lemmatization),
  inflect (English plural), pymorphy3 (Russian morphology), RuWordNet and
  NLTK WordNet (synonyms). Each engine is loaded lazily and degrades
  gracefully when absent.
- **Command-line interface** with configurable weights, seed, output path,
  and quiet/preview modes.
- **Honest diagnostics.** Active engines are reported to `stderr` before
  processing begins.

---

## Installation

### Minimum (stdlib only)

The script runs without any third-party packages, though with reduced
functionality:

```bash
git clone https://github.com/<your-user>/skavenizer.git
cd skavenizer
python skavenizer.py "Человек идёт домой"
```

Recommended

Install all optional engines for full functionality:

```bash
pip install python-docx pypdf textract ruwordnet nltk pymorphy3 pymystem3 inflect
python -c "import nltk; nltk.download('wordnet')"
```

On Windows, textract additionally requires antiword and libmagic for
.doc support. If these are unavailable, .doc inputs will fail with a
clear error message.

Python version

Python 3.9 or newer. The from __future__ import annotations import
ensures PEP 604 union syntax works on 3.9.

---

Usage

Basic

```bash
python skavenizer.py input.docx
python skavenizer.py "Человек идёт домой. The dog runs fast."
echo "some text" | python skavenizer.py -
```

Options

- `source` — Path to a file, raw text, or `-` for stdin.
- `-o, --output PATH` — Output path. Default: `<stem>_skavenized.txt` for files, `skavenized_output.txt` otherwise.
- `--seed INT` — Seed the RNG for reproducible output.
- `--weights DUP SYN TVA KEEP` — Four non-negative weights (default `0.15 0.15 0.15 0.55`).
- `-q, --quiet` — Suppress stdout output and diagnostics. Errors still go to stderr.
- `--preview N` — Print only the first N characters.

Examples

Reproducible run with a fixed seed:

```bash
python skavenizer.py input.txt --seed 42 -o out.txt
```

Only the "tvari" transformation (disable duplication, synonyms, and no-op):

```bash
python skavenizer.py input.txt --weights 0 0 1 0
```

Preview the first 300 characters:

```bash
python skavenizer.py input.pdf --preview 300
```

Batch mode (no stdout):

```bash
python skavenizer.py input.txt -q
```

---

Transformation Rules

Duplicate

The word is repeated with a hyphen. Punctuation and spacing are preserved.

```
да    → да-да
yes   → yes-yes
```

Synonym

A synonym is appended with a hyphen. The synonym is drawn from RuWordNet
(Russian) or NLTK WordNet (English). If no synonym is found, the word is
left unchanged.

```
хорошо → хорошо-впору
fast   → fast-quick
```

Tvari (beast)

Applied only to animate nouns.

Russian. The word is reduced to its root (via pymystem3 lemmatization
when available, otherwise a 4-letter heuristic), stripped of trailing vowels,
and suffixed with the connective -о- followed by тварь (singular) or
твари (plural). When the root ends in a vowel, the connective is dropped.

```
человек  (sg.) → человеко-тварь
люди     (pl.) → человеко-твари
кошка    (sg.) → кошко-тварь
кошки    (pl.) → кошко-твари
я        (sg.) → я-тварь
```

English. Only the plural ending is stripped; the root is never cut. The
suffix is -thing (singular) or -things (plural). Irregular plurals
(mice, children) and invariants (series, news, sheep) are handled
correctly when inflect is installed.

```
dog   (sg.) → dog-thing
dogs  (pl.) → dog-things
mouse (sg.) → mouse-thing
mice  (pl.) → mouse-things
series      → series-things
```

Keep

The word is emitted unchanged. This is the default outcome for the majority
of words (weight 0.55).

---

Optional Engines

- `pymystem3` — Russian lemmatization for accurate root extraction. Install: `pip install pymystem3`
- `inflect` — English irregular plural/singular handling. Install: `pip install inflect`
- `pymorphy3` — Russian animacy and number detection. Install: `pip install pymorphy3`
- `ruwordnet` — Russian synonyms. Install: `pip install ruwordnet`
- `nltk + WordNet` — English synonyms and animacy. Install: `pip install nltk` and `nltk.download('wordnet')`
- `pypdf` — PDF reading (falls back to PyPDF2). Install: `pip install pypdf`
- `python-docx` — `.docx` reading. Install: `pip install python-docx`
- `textract` — `.doc` reading. Install: `pip install textract`

When an engine is missing, the script prints a one-line warning to stderr
(before processing) and falls back to a heuristic. No exception is raised;
the output remains valid.

---

Function Reference

I/O

- `extract_text(path)` — Dispatches to the appropriate reader by file extension.
- `_read_text_file(path)` — Reads `.txt` with encoding fallback.
- `_decode_bytes(data)` — Decodes bytes as UTF-8 → CP1251 → Latin-1.

Morphology

- `is_animate_ru(word)` — Russian animacy via `pymorphy3`.
- `is_animate_en(word)` — English animacy via WordNet.
- `is_plural_ru(word)` — Russian number via `pymorphy3`.
- `is_plural_en(word)` — English number via `inflect` (fallback: suffix heuristic).
- `singular_en(word)` — Returns the singular form of an English word.
- `lemma_ru(word)` — Russian lemma via `pymystem3` (cached).
- `russian_root(word)` — Root extraction: `pymystem3` or heuristic.

Transformations

- `make_beast_ru(word)` — Russian "tvari" transformation.
- `make_beast_en(word, plural, sg)` — English "thing" transformation.
- `apply_case(source, target)` — Applies source's case pattern to target.
- `skavenize_word(word, weights)` — Random weighted transformation of a single word.
- `skavenize_text(text, weights)` — Applies `skavenize_word` to every token.

CLI

- `parse_args()` — Defines and validates CLI arguments.
- `_default_output_for(source, is_file)` — Chooses the default output path.
- `_print_diagnostics(quiet)` — Reports active engines to stderr.
- `main()` — Entry point.

---

Limitations

- PDF with Cyrillic. The `pypdf`/`PyPDF2` backends often produce garbage
  for Cyrillic PDFs. This is a library limitation, not a bug in Skavenizer.
- `.doc` on Windows. Requires `antiword` and `libmagic` installed and
  available on PATH.
- Irregular Russian roots. Without `pymystem3`, root extraction is a
  4-letter heuristic and can produce imperfect results on complex words.
- Synonym quality. RuWordNet and WordNet return the first available
  synonym, which may not always be idiomatic in context.
- Mixed-script tokens. A word is classified as Russian if it contains at
  least one Cyrillic character; tokens like `abcа` (Latin + Cyrillic `а`)
  will be treated as Russian.
- Irregular cases for `apply_case`. Mixed-case inputs (`СоБака`) are
  normalized to `Собак…`.
- Empty input. Exits with code 0 and prints `Text is empty.` to stderr.

---

Credits

This project was developed with the assistance of DeepSeek
(https://www.deepseek.com/), an AI language model used for design
consultation, code review, iterative refinement, and documentation.

The author retains full responsibility for the final implementation and
its behaviour.

---

License

MIT License. See LICENSE for details.