# Rippling New-Business Segmentation & Staffing (case study)

Final-round case study for Revenue Strategy & Operations. Built on the fictional dataset in `data/`.

## Deliverables
| File | What it is |
|---|---|
| `deck/Rippling_NB_Segmentation_Staffing.pptx` | Cover + 7-slide executive deck + appendix (A1–A6), with speaker notes |
| `model/Rippling_NB_Staffing_Model.xlsx` | Editable, formula-driven funnel analysis and staffing model (7 tabs; all checks pass) |
| `docs/Executive_Summary.docx` / `.md` | One-page written summary |
| `docs/Talk_Track_and_QA.md` | 15–20 minute talk track and 10 anticipated questions |

## Rebuild
```bash
python3 -I analysis/build_model.py data/rippling_nb_dataset.csv model/Rippling_NB_Staffing_Model.xlsx
python3 <xlsx-skill>/scripts/recalc.py model/Rippling_NB_Staffing_Model.xlsx   # computes formula values
python3 -I analysis/extract_deck_data.py model/Rippling_NB_Staffing_Model.xlsx deck/deck_data.json
NODE_PATH=<node_modules with pptxgenjs> node deck/build_deck.js deck/deck_data.json deck/Rippling_NB_Segmentation_Staffing.pptx <pptx-skill>/scripts/apply_theme.js
```
Deck numbers are read from the recalculated workbook, so the slides and the model reconcile.

`analysis/phase1_explore.py` and `phase1_extra.py` hold the initial pandas exploration.
