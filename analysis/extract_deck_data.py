"""Read recalculated workbook values into JSON so the deck reconciles to the model."""
import json, sys
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
f, s, c = wb["Funnel"], wb["Staffing"], wb["Cost_Optional"]
seg = ["C", "D", "E", "F"]
fcols = dict(zip(["leads","demos","qo","logos","arr","l2d","d2o","o2w","l2w","asp","arr_lead","lead_share","arr_share","demos_logo"], "BCDEFGHIJKLMNO"))
def frow(r): return {k: f[f"{v}{r}"].value for k, v in fcols.items()}
out = {
  "segments": [f[f"A{r}"].value for r in range(6, 10)],
  "seg_funnel": [frow(r) for r in range(6, 11)],
  "bucket_names": [f[f"A{r}"].value for r in range(14, 19)],
  "bucket_funnel": [frow(r) for r in range(14, 20)],
  "monthly": [{"month": f[f"A{r}"].value.strftime("%b-%y"), "leads": f[f"B{r}"].value, "demos": f[f"C{r}"].value,
               "logos": f[f"D{r}"].value, "arr": f[f"E{r}"].value, "l2d": f[f"F{r}"].value, "mix101": f[f"I{r}"].value} for r in range(23, 35)],
  "early_vs_base": {lab: [f[f"{col}{r}"].value for col in "BCDEF"] for lab, r in
      [("el",41),("bl",42),("gl",43),("ea",44),("ba",45),("ga",46),("ep",47),("bp",48),("gp",49),("ed",50),("bd",51),("cd",52),
       ("eq",53),("bq",54),("cq",55),("ew",56),("bw",57),("cw",58)]},
  "mix101": [f["B59"].value, f["C59"].value],
  "staff": {lab: [s[f"{col}{r}"].value for col in seg + ["G"]] for lab, r in
      [("leads",7),("demos",8),("l2d",9),("d2w",10),("asp",11),("sdr_cap",14),("ae_cap",15),("sc_rate",16),("sc_cap",17),("util",18),
       ("f_sdr",21),("f_ae",22),("f_sc",23),("h_sdr",26),("h_ae",27),("ae_per_sdr",30),("ae_util",32),("arr_m",35),("arr_y",36),("arr_ae",37),
       ("e_l2d",40),("u_l2d",42),("u_demos",43),("u_f_ae",44),("u_f_sc",45),("u_h_sdr",46),("u_h_ae",47),("u_arr_y",50),("d_arr",51),("d_ae",52)]},
  "sc_pool": s["C28"].value, "sc_ent": s["F28"].value, "sc_tot": s["G28"].value, "hc_tot": s["G29"].value,
  "u_sc_pool": s["C48"].value, "u_sc_ent": s["F48"].value, "u_sc_tot": s["G48"].value, "u_hc_tot": s["G49"].value, "d_hc_tot": s["G53"].value,
  "ae_per_sc_pool": s["C31"].value, "ae_per_sc_ent": s["F31"].value,
  "cost": {"unit": [c[f"C{r}"].value for r in (5,6,7)], "base": [c[f"E{r}"].value for r in (5,6,7,8)], "up": [c[f"G{r}"].value for r in (5,6,7,8)]},
  "checks_pass": wb["Checks"]["E22"].value,
}
json.dump(out, open(sys.argv[2], "w"), indent=1, default=str)
print("checks_pass", out["checks_pass"], "hc", out["hc_tot"], out["u_hc_tot"])
