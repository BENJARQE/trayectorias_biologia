# Biology EAFIT graduate trajectories

Static HTML, CSS, vanilla JavaScript and SVG. No framework or web build step. Prepared locally; publication requires the owner's approval.

The visual shell, stylesheet and SVG components reuse [Jorge Arango's Music trajectories](https://github.com/JarangoM/trayectorias_musica), with express permission reported by the project owner. Reference commit: `a0c276356c3852c2f6cf971ed4f52fa2d6aa69f7`. The interface accent is `#70B77E` (dark) and `#237A43` (light). The categorical charts now use an expanded, append-only color registry, selected by perceptual distance, with unique colors for each label in both themes. Attribution is retained here and in source comments, without visible page credits.

## Local preview

From this directory: `python -m http.server 8765 --bind 127.0.0.1`. Open [the local preview](http://127.0.0.1:8765). JSON fetch requires HTTP, rather than opening an HTML file directly.

## Reproduce aggregates locally

Python 3.11+ and `openpyxl` (read-only Excel extraction). `_build_data.py` takes explicit local inputs; no personal source files belong in this repository:

```text
python _build_data.py --review PATH_TO_CORRECTED_REVIEW_JSON --employers PATH_TO_EMPLOYER_AUDIT_XLSX --sizes PATH_TO_SIZE_AUDIT_XLSX --practice PATH_TO_CLOSED_PRACTICES_JSON --comfama PATH_TO_CLOSED_COMFAMA_JSON --private-output DIRECTORY_OUTSIDE_THIS_REPO --out datos/biologia_datos.json
```

The research repository owns the definitive closure, frozen input copies, private review, pending size register and verification records. Presentation revisions read these as immutable inputs and write new derived records to a separate external directory. The generator refuses a private output location inside the public repository. The closure workbooks, original titles, professional classifications, employer mappings and selected positions are not edited.

The release includes 96 precomputed filter views, two counting units, two longitudinal sample modes and two size-confidence modes. Filters do not recompute microdata in the browser. Original positions and the five-or-six-experience sample are fixed before filtering. Comfama episodes never enter LinkedIn positions. Employer sector and professional area remain independent. Pending audit status and all source uncertainty remain in the private registry.

## Disclosure controls and limits

Only aggregates are served. Named categories, institutional employers and normalized titles require at least two distinct graduates in the applicable view. Remaining records, independent professionals and natural-person employers are pooled without dropping observations. Original titles remain private; the display mapping merges only clear language/gender equivalents and retains specialisms and seniority. `Biólogo` is the generic display term.

Counts are exact integers; means retain their calculated precision in JSON and display up to two decimals. No rounding to five or hundreds is performed. Unique-graduate charts use bars, with percentages over the distinct people represented, rather than shares of summed memberships. Memberships can overlap.

The four all-cohort reference source/postgraduate scenarios are exact aggregate releases: 323, 304, 380 and 361 observations. Narrower cohort/combined-filter comparisons in a metric family are restricted if differencing would isolate a single known contributor. If any cell of a partition is unsafe, the entire related distribution is withheld; size components are protected together to prevent recovery by subtraction. Named rows still require two graduates, including in reference views. Suppressed values remain null and are distinct from zero, missing information or exclusions. These controls reduce disclosure risk and do not claim formal differential privacy or protection against arbitrary external auxiliary information.

Colors are keyed by the complete displayed label, not a hash or a seven-color cycle. Selection maximizes perceptual separation between labels that co-occur in charts, rather than spending distinct hues on unrelated chart families. Existing labels keep their assigned pair when filters or views change. The public payload contains only eligible labels and anonymous grouped remainders, never identities or record keys. Coverage includes four LinkedIn/Comfama combinations plus document-based reasons for absent selected LinkedIn experience. Mean-only longitudinal size charts preserve the six original positions, denominators, missing-size counts and filter exclusions.

## Publication after approval

Target: `BENJARQE/trayectorias_biologia`. Initialize this directory independently of the research repository. Only allowlisted site files, documentation and the generic generator may be committed. Never publish the research parent directory. The prepared GitHub Pages workflow deploys only allowlisted HTML/CSS/JS/JSON assets from this repository. No upload, repository creation or push is performed until approval.

GitHub CLI must authenticate locally with permission to create/push the target repository. Use `gh auth status`, then `gh auth login --hostname github.com --web` only if necessary. The research handoff contains the exact preparation commands. No credentials are stored in this repository.
