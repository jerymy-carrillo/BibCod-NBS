# BibCod-NBS

**Bibliographic Coding Tool for Nature-Based Solutions Systematic Reviews**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.12](https://img.shields.io/badge/python-3.12-blue.svg)](https://www.python.org/downloads/)

<!-- After connecting the repository to Zenodo (see "How to cite"),
     paste the DOI badge provided by Zenodo on the line below: -->
<!-- [![DOI](https://zenodo.org/badge/1279700188.svg)](https://doi.org/10.5281/zenodo.20837764) -->

---

## Overview

**BibCod-NBS** is a Python tool developed *ad hoc* to support a systematic
mapping review on **Nature-Based Solutions (NBS) applied to water management**,
conducted within the framework of the **CYTED Network** (Ibero-American
Programme of Science and Technology for Development) and led by the Department
of Geology, **University of Oviedo** (Spain).

The tool processes a bibliographic corpus exported from Scopus and supports
the structured extraction, coding, and thematic scoring of scientific articles
according to a predefined 43-variable protocol.

---

## What the tool does

Given a bibliographic file in **RIS format** (`.ris`) exported from Scopus,
the script:

1. **Parses** the file and extracts the metadata of each record
   (authors, title, year, journal, DOI).
2. **Provides** the full structured coding protocol — **10 thematic blocks**
   and **43 variables** (dichotomous, categorical, and ordinal) — used by the
   research team to code each article after full-text reading.
3. **Computes** a thematic affinity score (the **DS Index**, scale 1–10) for
   each article, measuring its alignment with the journal collection
   *"Nature-based Solutions applied to water management"*
   (*Discover Sustainability*, Springer).
4. **Exports** a fully formatted Excel workbook containing the coded database,
   a prioritised sub-corpus, descriptive statistics, and the coding protocol.

---

## Methodological note

The 43 coding variables require **expert judgement based on full-text reading**
of each article. This tool does **not** automate that judgement: it automates
the **data structure**, the **DS Index calculation**, and the
**reporting/export** steps.

In the original review (n = 85 articles), the qualitative coding was carried
out **independently by at least two members of the research team**, with
discrepancies resolved by consensus, following good-practice guidance for data
extraction forms in systematic reviews (Büchter et al., 2020) and
semi-automated extraction workflows (Schmidt et al., 2025).

---

## Requirements

- Python 3.12 or higher
- [`openpyxl`](https://openpyxl.readthedocs.io/) ≥ 3.1.0

Install the dependency with:

```bash
pip install openpyxl
```

---

## Usage

```bash
python bibcod_nbs.py --input bibliography.ris --output results.xlsx
```

### Command-line arguments

| Argument | Short | Description | Default |
|----------|-------|-------------|---------|
| `--input` | `-i` | Path to the RIS file exported from Scopus | `bibliography.ris` |
| `--output` | `-o` | Path to the Excel file to be generated | `bibcod_nbs_results.xlsx` |

### Example

```bash
python bibcod_nbs.py -i my_corpus.ris -o coded_database.xlsx
```

---

## Output structure

The generated Excel workbook contains four sheets:

| Sheet | Content |
|-------|---------|
| **Full Analysis** | Complete coded database (articles × 43 variables + DS Index) |
| **TOP Articles DS≥7** | Prioritised sub-corpus ranked by DS Index |
| **Descriptive Stats** | Frequencies and percentages by thematic block |
| **Coding Protocol** | Full protocol: 10 blocks, 43 variables |

---

## The coding protocol (10 blocks, 43 variables)

1. **Publication metadata** — year, journal, quartile
2. **Article type and intervention** — type, intervention nature, case study
3. **Methodological rigour** — clarity, novelty, data quality, limitations
4. **Geographic and social context** — country, zone, community, vulnerability, scale
5. **Hydrology and geological discipline** — discipline, water system, integrated management
6. **NBS typology** — SUDS, river/catchment restoration, wetlands, reforestation, etc.
7. **Water-related problem** — floods, droughts, water quality, aquifer recharge
8. **Climate change** — driver, microclimate, monitoring
9. **Planning and regulatory framework** — planning, instruments, conflicts, policy, grey comparison
10. **Social dimension** — participation, acceptance, traditional knowledge, applicability, transferability

The full set of operative questions and admissible values for each variable
is documented inside the script (`PROTOCOL_BLOCKS`) and exported to the
"Coding Protocol" sheet of the output workbook.

---

## How to cite

If you use this tool in your research, please cite it as follows.

<!-- Once the Zenodo DOI is generated, replace [DOI] below with the actual DOI. -->

> [Jerymy Antonio Carrillo], [J.A., Carrillo]. (2026). *BibCod-NBS: Bibliographic Coding
> Tool for Nature-Based Solutions Systematic Reviews* (Version 1.0)
> [Computer software]. https://doi.org/10.5281/zenodo.20837764

**Authors / Contributors:**
- Jerymy Antonio Carrillo Bravo — ORCID: 0000-0002-3059-0609
- Research team, CYTED Network / University of Oviedo
- Principal Investigator: Dr. María José Domínguez-Cuesta (University of Oviedo)

---

## References

- Büchter, R. B., Weise, A., & Pieper, D. (2020). Development, testing and use
  of data extraction forms in systematic reviews: a review of methodological
  guidance. *BMC Medical Research Methodology*, 20, 259.
  https://doi.org/10.1186/s12874-020-01143-3

- Page, M. J., McKenzie, J. E., Bossuyt, P. M., et al. (2021). The PRISMA 2020
  statement: an updated guideline for reporting systematic reviews. *BMJ*,
  372, n71. https://doi.org/10.1136/bmj.n71

- Schmidt, L., et al. (2025). Data extraction methods for systematic review
  (semi)automation: update of a living systematic review [version 3].
  *F1000Research*, 10, 401. https://doi.org/10.12688/f1000research.51117.3

---

## License

This project is licensed under the **MIT License** — see the
[LICENSE](LICENSE) file for details.

---

## Acknowledgements

This work was developed within the framework of the **CYTED Network**
(Ibero-American Programme of Science and Technology for Development),
Department of Geology, **University of Oviedo** (Spain).
