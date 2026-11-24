# Biology EAFIT graduate trajectories

Static HTML, CSS, vanilla JavaScript and SVG. No framework or web build step. Prepared locally; publication requires the owner's approval.

The visual shell, stylesheet, SVG components and chart palette reuse [Jorge Arango's Music trajectories](https://github.com/JarangoM/trayectorias_musica), with express permission reported by the project owner. Reference commit: `a0c276356c3852c2f6cf971ed4f52fa2d6aa69f7`. The two interface accent declarations are changed to `#70B77E` (dark) and `#237A43` (light); the original chart colors remain.

## Local preview

From this directory: `python -m http.server 8765 --bind 127.0.0.1`. Open [the local preview](http://127.0.0.1:8765). JSON fetch requires HTTP, rather than opening an HTML file directly.

## Reproduce aggregates locally

Python 3.11+ and `openpyxl` (read-only Excel extraction). `_build_data.py` takes explicit local inputs; no personal source files belong in this repository:

```text
python _build_data.py --review PATH_TO_CORRECTED_REVIEW_JSON --employers PATH_TO_EMPLOYER_AUDIT_XLSX --sizes PATH_TO_SIZE_AUDIT_XLSX --practice PATH_TO_CLOSED_PRACTICES_JSON --comfama PATH_TO_CLOSED_COMFAMA_JSON --private-output DIRECTORY_OUTSIDE_THIS_REPO --out datos/biologia_datos.json
```

The research repository owns the persistent selection correction, frozen input copies, source hashes, private enriched review, exact statistics, pending size register and integrated verification report. This public generator exports both exact private outputs to the external directory and a separate protected aggregate payload here. It refuses a private output location inside this public directory.

The release includes 96 precomputed filter views, two counting units, two longitudinal sample modes and two size-confidence modes. Filters do not recompute microdata in the browser. Original positions and the five-or-six-experience sample are fixed before filtering. Comfama episodes never enter LinkedIn positions. Employer sector and professional area remain independent. Pending audit status and all source uncertainty remain in the private registry.

## Disclosure controls and limits

Only aggregates are served. At least five distinct graduates support a named cell; complementary suppression and a guard over the entire released query lattice restrict small differences between filter views. Counts are rounded to base five, and size statistics require at least ten graduates. The generator omits small category labels and employers recorded as natural persons. Suppressed values remain null, not zero. The exact overall closure totals are explicitly authorized controls. These controls reduce disclosure risk; they do not claim formal differential privacy. See the site's methodology for interpretation, denominators and source limits.

## Publication after approval

Target: `BENJARQE/trayectorias_biologia`. Initialize this directory independently of the research repository. Only allowlisted site files, documentation and the generic generator may be committed. Never publish the research parent directory. The prepared GitHub Pages workflow deploys only allowlisted HTML/CSS/JS/JSON assets from this repository. No upload, repository creation or push is performed until approval.

GitHub CLI must authenticate locally with permission to create/push the target repository. Use `gh auth status`, then `gh auth login --hostname github.com --web` only if necessary. The research handoff contains the exact preparation commands. No credentials are stored in this repository.
