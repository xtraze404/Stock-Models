"""Turn each company's ClinicalTrials.gov studies into programs (one per indication bucket) with stage, approval odds,
peak sales, launch year and next readout.  usage: parse_pipeline.py [TICKERS...]  -> programs/<T>.json + summary
Approval odds: BIO / Informa / QLS, Clinical Development Success Rates 2011-2020 (Feb 2021), phase-transition rates by
disease area; non-oncology rare diseases and chronic high-prevalence diseases use the report's own rates."""
import datetime as dt, json, os, re, sys
from collections import Counter, defaultdict

D = os.path.dirname(os.path.abspath(__file__))
TODAY = dt.date(2026, 10, 21)
YEAR0 = 2026

# phase transitions: P1->P2, P2->P3, P3->NDA/BLA, NDA/BLA->approval  (BIO 2011-2020, Fig. 2 and Fig. 8b)
RATES = {
    "Hematology": (0.696, 0.481, 0.768, 0.931), "Metabolic": (0.618, 0.450, 0.636, 0.875),
    "Infectious disease": (0.578, 0.384, 0.640, 0.929), "Other": (0.636, 0.386, 0.600, 0.884),
    "Ophthalmology": (0.716, 0.355, 0.512, 0.911), "Autoimmune": (0.552, 0.314, 0.653, 0.941),
    "Allergy": (0.564, 0.283, 0.647, 1.000), "Gastroenterology": (0.467, 0.342, 0.571, 0.909),
    "Respiratory": (0.559, 0.219, 0.645, 0.956), "Psychiatry": (0.527, 0.268, 0.563, 0.912),
    "Endocrine": (0.433, 0.266, 0.662, 0.863), "Neurology": (0.477, 0.268, 0.531, 0.867),
    "Oncology": (0.488, 0.246, 0.477, 0.920), "Cardiovascular": (0.500, 0.210, 0.552, 0.825),
    "Urology": (0.409, 0.150, 0.692, 0.846),
    "Rare (non-oncology)": (0.674, 0.446, 0.604, 0.936), "Chronic high-prevalence": (0.460, 0.231, 0.595, 0.926),
}

# indication buckets: (bucket, keywords, BIO area, rate class override, peak $M = average peak of a successful launch in the indication)
# rate class: "rare" -> Rare (non-oncology), "chronic" -> Chronic high-prevalence, None -> the area's own rates
B = [
    # oncology
    ("Pancreatic cancer", r"pancrea|pdac", "Oncology", None, 2500),
    ("Lung cancer", r"lung cancer|lung carcinoma|nsclc|sclc|non-small cell|small cell lung|lung adenocarcinoma|mesothelioma", "Oncology", None, 3000),
    ("Colorectal cancer", r"colorectal|colon cancer|rectal cancer|\bcrc\b|colon adenocarcinoma", "Oncology", None, 2000),
    ("Breast cancer", r"breast", "Oncology", None, 3000),
    ("Prostate cancer", r"prostat(e|ic) (cancer|carcinoma|adenocarcinoma)|mcrpc|crpc", "Oncology", None, 2500),
    ("Melanoma", r"melanoma", "Oncology", None, 1500),
    ("Ovarian cancer", r"ovarian|fallopian|peritoneal (cancer|carcinoma)", "Oncology", None, 1200),
    ("Bladder / urothelial cancer", r"bladder cancer|urothelial|nmibc", "Oncology", None, 1200),
    ("Kidney cancer", r"renal cell|kidney cancer|\brcc\b|clear cell", "Oncology", None, 1500),
    ("Liver cancer", r"hepatocellular|\bhcc\b|liver cancer|cholangiocarcinoma|biliary tract cancer|biliary cancer", "Oncology", None, 1200),
    ("Gastric / esophageal cancer", r"gastric|stomach cancer|esophag|gastroesophageal|\bgej\b", "Oncology", None, 1200),
    ("Head and neck cancer", r"head and neck|hnscc|oral cancer|nasopharyn|laryn", "Oncology", None, 1000),
    ("Brain cancer", r"glioblastoma|glioma|\bgbm\b|brain (tumor|cancer|metast)|astrocytoma", "Oncology", None, 800),
    ("Gynecologic cancer", r"endometrial|cervical cancer|uterine", "Oncology", None, 1000),
    ("Sarcoma / GIST", r"sarcoma|\bgist\b|gastrointestinal stromal|desmoid", "Oncology", None, 600),
    ("Thyroid / neuroendocrine cancer", r"thyroid cancer|neuroendocrine|carcinoid|pheochromocytoma", "Oncology", None, 600),
    ("Multiple myeloma", r"myeloma|amyloidosis al|al amyloidosis", "Oncology", None, 2500),
    ("Leukemia", r"leukemia|leukaemia|\baml\b|acute lymphoblastic|\bcll\b|\bcml\b", "Oncology", None, 1500),
    ("Lymphoma", r"lymphoma|dlbcl|hodgkin|mantle cell|follicular", "Oncology", None, 1500),
    ("Myelofibrosis / MDS / MPN", r"myelofibrosis|myelodysplastic|\bmds\b|polycythemia|thrombocythemia|myeloproliferative|mastocytosis", "Oncology", None, 1000),
    ("Solid tumors (basket)", r"solid tumou?r|advanced cancer|metastatic cancer|neoplasm|carcinoma|cancer|tumor|tumour|malignan", "Oncology", None, 1500),
    # metabolic / endocrine / cardio
    ("Obesity", r"obesity|overweight|weight (loss|management)", "Metabolic", "chronic", 6000),
    ("Type 2 diabetes", r"type 2 diabetes|t2d|diabetes mellitus, type 2|diabetes$|\bdiabetes\b(?!.*type 1)", "Metabolic", "chronic", 2500),
    ("Type 1 diabetes", r"type 1 diabetes|t1d", "Metabolic", None, 1000),
    ("MASH / liver fibrosis", r"nash|mash|steatohepatitis|steatotic|nafld|masld|liver fibrosis|cirrhosis", "Metabolic", "chronic", 2500),
    ("Lipids / cardiometabolic", r"cholesterol|lipoprotein|lp\(a\)|hyperlipid|dyslipid|triglycerid|ldl", "Cardiovascular", "chronic", 3000),
    ("Heart failure", r"heart failure|hfref|hfpef|cardiac failure", "Cardiovascular", "chronic", 2500),
    ("Cardiomyopathy", r"cardiomyopathy|\bhcm\b|myocarditis", "Cardiovascular", None, 2000),
    ("ATTR amyloidosis", r"transthyretin|attr", "Cardiovascular", "rare", 3000),
    ("Pulmonary hypertension", r"pulmonary (arterial )?hypertension|\bpah\b|ph-ild", "Cardiovascular", "rare", 1500),
    ("Hypertension", r"hypertension(?!.*pulmonary)|blood pressure", "Cardiovascular", "chronic", 2000),
    ("Thrombosis / cardiovascular", r"thrombosis|thromboembol|venous thrombo|embolism|atrial fibrillation|stroke|coronary|myocardial infarction|cardiovascular|arrhythm", "Cardiovascular", None, 1500),
    ("Endocrine (rare)", r"cushing|acromegaly|hypoparathyroid|congenital adrenal|achondroplasia|growth hormone|hypophosphat|hyperinsulin|prader", "Endocrine", "rare", 700),
    ("Rare metabolic / lysosomal", r"phenylketonuria|\bpku\b|fabry|gaucher|pompe|mucopolysacchar|homocystinuria|urea cycle|wilson|glycogen storage|lipodystrophy|hypophosphatasia|metabolic disorder|organic acid", "Metabolic", "rare", 700),
    # CNS
    ("Alzheimer's disease", r"alzheimer|dementia|cognitive impairment", "Neurology", "chronic", 4000),
    ("Parkinson's disease", r"parkinson", "Neurology", None, 1500),
    ("ALS", r"amyotrophic|\bals\b|motor neuron", "Neurology", "rare", 800),
    ("Huntington's disease", r"huntington", "Neurology", "rare", 800),
    ("Multiple sclerosis", r"multiple sclerosis|\bms\b", "Neurology", None, 2000),
    ("Rare epilepsy", r"dravet|lennox|developmental and epileptic|\bdee\b|scn2a|scn8a|cdkl5|tuberous sclerosis|infantile spasm", "Neurology", "rare", 800),
    ("Epilepsy", r"epilep|seizure", "Neurology", None, 1200),
    ("Migraine / pain", r"migraine|headache|pain|neuropath|fibromyalgia", "Neurology", None, 1200),
    ("Muscular dystrophy / SMA", r"duchenne|muscular dystrophy|\bdmd\b|spinal muscular|\bsma\b|myotonic|facioscapulohumeral|limb.girdle|friedreich|ataxia", "Neurology", "rare", 1000),
    ("Myasthenia gravis / neuromuscular", r"myasthenia|cidp|neuromyelitis|guillain", "Neurology", "rare", 1500),
    ("Rare neurodevelopmental", r"rett|angelman|fragile x|phelan|prader-willi|batten|leukodystrophy|niemann", "Neurology", "rare", 600),
    ("Movement / sleep disorders", r"tremor|dyskinesia|narcolepsy|insomnia|sleep|restless legs|tourette|dystonia", "Neurology", None, 1200),
    ("Schizophrenia / psychosis", r"schizophren|psychos", "Psychiatry", "chronic", 2500),
    ("Depression / mood", r"depress|bipolar|\bmdd\b|mood", "Psychiatry", "chronic", 2000),
    ("Anxiety / PTSD / addiction", r"anxiety|ptsd|post-traumatic|addiction|use disorder|opioid|alcohol|smoking|cocaine|autism|adhd|attention deficit", "Psychiatry", None, 800),
    # immunology / inflammation
    ("Atopic dermatitis", r"atopic dermatitis|eczema", "Autoimmune", None, 3000),
    ("Psoriasis", r"psoria", "Autoimmune", None, 2500),
    ("IBD (UC / Crohn's)", r"ulcerative colitis|crohn|inflammatory bowel", "Autoimmune", None, 3000),
    ("Rheumatoid / spondyloarthritis", r"rheumatoid|spondyl|psoriatic arthritis|arthritis", "Autoimmune", None, 2000),
    ("Lupus / Sjogren's", r"lupus|sjogren|sjögren", "Autoimmune", None, 1500),
    ("Hidradenitis / alopecia / vitiligo", r"hidradenitis|alopecia|vitiligo|urticaria|prurigo|pemphig", "Autoimmune", None, 1500),
    ("Rare autoimmune", r"myositis|vasculitis|scleroderma|systemic sclerosis|igg4|thyroid eye|graves|immune thrombocytopenia|\bitp\b|autoimmune hemolytic|antiphospholipid|igan|iga nephropathy", "Autoimmune", "rare", 800),
    ("Asthma / COPD", r"asthma|copd|chronic obstructive", "Respiratory", "chronic", 2500),
    ("Eosinophilic / allergy", r"eosinophil|allerg|food allergy|anaphyla|rhinitis|rhinosinusitis|nasal polyp", "Allergy", None, 1000),
    ("Celiac / GI", r"celiac|gastroparesis|irritable bowel|\bibs\b|short bowel|pancreatitis|constipation|dyspepsia", "Gastroenterology", None, 800),
    # respiratory / renal / liver
    ("Idiopathic pulmonary fibrosis", r"pulmonary fibrosis|\bipf\b|interstitial lung", "Respiratory", None, 2000),
    ("Cystic fibrosis / bronchiectasis", r"cystic fibrosis|bronchiectasis|alpha-1|alpha1|antitrypsin", "Respiratory", "rare", 1500),
    ("Kidney disease", r"kidney|renal|nephr|fsgs|polycystic|adpkd|apol1|glomerul|dialysis|hyperkal", "Other", None, 1200),
    ("Cholestatic liver disease", r"biliary cholangitis|\bpbc\b|sclerosing cholangitis|\bpsc\b|cholesta|alagille|hepatic encephal", "Gastroenterology", "rare", 800),
    # hematology
    ("Sickle cell / thalassemia", r"sickle|thalass", "Hematology", "rare", 1000),
    ("Hemophilia / bleeding", r"hemophilia|haemophilia|von willebrand|bleeding|coagulation factor", "Hematology", "rare", 1000),
    ("Rare hematology", r"paroxysmal|\bpnh\b|anemia|anaemia|neutropenia|aplastic|hemoly|iron overload|polycyth", "Hematology", None, 700),
    ("Hereditary angioedema", r"angioedema|\bhae\b", "Hematology", "rare", 1200),
    # infectious / vaccines
    ("Pneumococcal vaccine", r"pneumococ", "Infectious disease", None, 6000),
    ("Vaccines / respiratory viruses", r"vaccin|influenza|\brsv\b|respiratory syncytial|covid|sars-cov|coronavirus", "Infectious disease", None, 1000),
    ("HIV / hepatitis", r"\bhiv\b|hepatitis b|hepatitis d|\bhbv\b|\bhdv\b|hepatitis c", "Infectious disease", None, 1500),
    ("Bacterial / fungal infections", r"infection|bacter|fungal|candid|aspergill|pneumonia|sepsis|c\. difficile|clostridi|tuberculosis|uti\b|urinary tract inf", "Infectious disease", None, 500),
    ("Viral infections (other)", r"\bcmv\b|cytomegalovirus|herpes|hpv|papilloma|ebv|epstein|adenovirus|dengue|norovirus", "Infectious disease", None, 600),
    # ophthalmology
    ("Retinal disease (AMD / DME)", r"macular|\bamd\b|diabetic retinopathy|\bdme\b|retinal vein|choroidal|neovascular", "Ophthalmology", None, 2500),
    ("Geographic atrophy", r"geographic atrophy", "Ophthalmology", None, 1500),
    ("Inherited retinal disease", r"retinitis pigmentosa|inherited retinal|leber|stargardt|choroideremia|achromatopsia|x-linked retinoschisis|usher", "Ophthalmology", "rare", 500),
    ("Eye disease (other)", r"glaucoma|dry eye|presbyopia|myopia|uveitis|keratitis|conjunctivitis|ocular|blephar|cornea|cataract|thyroid eye", "Ophthalmology", None, 600),
    # other
    ("Women's health", r"endometriosis|fibroid|menopaus|vasomotor|contracept|preeclampsia|preterm|polycystic ovary|infertility", "Endocrine", None, 1000),
    ("Urology", r"overactive bladder|incontinence|prostatic hyperplasia|\bbph\b|interstitial cystitis|erectile", "Urology", None, 600),
    ("Dermatology (other)", r"acne|rosacea|epidermolysis|ichthyosis|wound|scar|burn|skin", "Other", None, 500),
    ("Musculoskeletal", r"osteoarthritis|osteoporosis|bone|tendon|fracture|sarcopenia|cachexia", "Other", None, 800),
    ("Transplant / GVHD", r"transplant|graft.versus|gvhd|rejection", "Autoimmune", None, 700),
]
B = [(n, re.compile(k, re.I), a, rc, pk) for n, k, a, rc, pk in B]
SKIP_COND = re.compile(r"^healthy|healthy (volunteer|subject|participant|adult)|^normal volunteer|pharmacokinetic|bioequivalence|drug.drug interaction|hepatic impairment|renal impairment", re.I)
SOC = re.compile(r"placebo|saline|vehicle|standard of care|best supportive|sham|docetaxel|paclitaxel|nab-paclitaxel|carboplatin|cisplatin|gemcitabine|pemetrexed|oxaliplatin|irinotecan|fluorouracil|5-fu|leucovorin|capecitabine|folfirinox|folfox|folfiri|cetuximab|bevacizumab|pembrolizumab|nivolumab|atezolizumab|durvalumab|ipilimumab|trastuzumab|rituximab|dexamethasone|prednisone|lenalidomide|bortezomib|azacitidine|venetoclax|letrozole|fulvestrant|enzalutamide|abiraterone|metformin|semaglutide|insulin|midazolam|itraconazole|rifampin|caffeine|moxifloxacin|chemotherapy|radiation|radiotherapy|surgery|regimen|investigator.s choice|physician.s choice", re.I)
PARTNER_CLASSES = {"INDUSTRY"}
PH = {"EARLY_PHASE1": 1.0, "PHASE1": 1.0, "PHASE2": 2.0, "PHASE3": 3.0, "PHASE4": 4.0}
LAUNCH_LAG = {4.0: 0, 3.5: 1, 3.0: 3, 2.5: 4, 2.0: 5, 1.5: 6, 1.0: 7}
OWNED_MARGIN, ROYALTY = 0.40, 0.12
MAX_PROGRAMS = 8
# biomarker-selected oncology niches get 40% of the bucket peak (pan-RAS / broad mutant populations excluded)
NICHE = re.compile(r"exon 20|exon20|\balk\b|alk-positive|\bros1\b|\bret\b|ret fusion|ntrk|met exon 14|\bfgfr|kras g12c|g12c|"
                   r"braf v600|v600e|\bidh[12]?\b|flt3|npm1|her2-mutant|her2 mutant|\bmsi-h|nrg1|claudin|cldn18|pik3ca", re.I)
BROAD = re.compile(r"pan-ras|ras\(on\)|ras-mutant|ras mutant|kras-mutant|kras mutant|kras g12x|nras|hras", re.I)
EXTRA_RAW = {"ROIV": ["Priovant", "Pulmovant"]}


def stage_of(phases):
    ps = sorted(PH[p] for p in (phases or []) if p in PH)
    if not ps:
        return None
    if len(ps) >= 2 and ps[-1] == ps[0] + 1 and ps[-1] in (2.0, 3.0):   # Phase 1/2 -> 1.5, Phase 2/3 -> 2.5
        return ps[0] + 0.5
    return ps[-1]


def loa(stage, area, rclass):
    r = RATES["Rare (non-oncology)" if rclass == "rare" and area != "Oncology" else
              "Chronic high-prevalence" if rclass == "chronic" else area]
    full = {1.0: r[0] * r[1] * r[2] * r[3], 2.0: r[1] * r[2] * r[3], 3.0: r[2] * r[3], 3.5: r[3], 4.0: 1.0}
    if stage in full:
        return full[stage]
    lo, hi = {1.5: (1.0, 2.0), 2.5: (2.0, 3.0)}[stage]
    return (full[lo] * full[hi]) ** 0.5


def bucket_of(conds):
    hits = Counter()
    for c in conds:
        if SKIP_COND.search(c):
            continue
        for n, k, a, rc, pk in B:
            if k.search(c):
                hits[n] += 1
                break
    return hits.most_common(1)[0][0] if hits else None


def parse(t, raw, name):
    key = raw["query"].lower()
    alias_keys = [a.lower() for a in EXTRA_RAW.get(t, [])]
    key_tokens = [w for w in re.findall(r"[a-z0-9]+", key) if len(w) > 2][:2]
    match = lambda s: bool(s) and (key in s.lower() or any(a in s.lower() for a in alias_keys))
    prog = defaultdict(lambda: {"trials": [], "drugs": Counter(), "owned": False, "partners": set(), "conds": []})
    kept = 0
    for s in raw["studies"]:
        p = s.get("protocolSection", {})
        if p.get("designModule", {}).get("studyType") not in (None, "INTERVENTIONAL"):
            continue
        st = p.get("statusModule", {})
        status = st.get("overallStatus")
        if status in ("WITHDRAWN", "TERMINATED", "SUSPENDED", "NO_LONGER_AVAILABLE", "APPROVED_FOR_MARKETING",
                      "AVAILABLE", "TEMPORARILY_NOT_AVAILABLE"):
            continue
        pcd = (st.get("primaryCompletionDateStruct") or {}).get("date") or ""
        upd = (st.get("lastUpdatePostDateStruct") or {}).get("date") or ""
        if status == "COMPLETED" and pcd[:4] < "2021":
            continue
        if status == "UNKNOWN" and upd[:4] < "2023":
            continue
        sp = p.get("sponsorCollaboratorsModule", {})
        lead = sp.get("leadSponsor", {})
        collabs = [c.get("name", "") for c in sp.get("collaborators", [])]
        if match(lead.get("name")):
            owned = True
        elif any(match(c) for c in collabs) and lead.get("class") in PARTNER_CLASSES:
            owned = False
        else:
            continue
        stg = stage_of(p.get("designModule", {}).get("phases"))
        if stg is None:
            continue
        b = bucket_of(p.get("conditionsModule", {}).get("conditions", []))
        if not b:
            continue
        drugs = [i["name"] for i in p.get("armsInterventionsModule", {}).get("interventions", [])
                 if i.get("type") in ("DRUG", "BIOLOGICAL", "GENETIC", "COMBINATION_PRODUCT") and not SOC.search(i.get("name", ""))]
        g = prog[b]
        g["conds"] += p.get("conditionsModule", {}).get("conditions", [])[:4]
        g["trials"].append({"nct": p["identificationModule"]["nctId"], "stage": stg, "status": status, "pcd": pcd,
                            "title": p["identificationModule"].get("briefTitle", "")[:120], "owned": owned})
        for d_ in drugs[:2]:
            g["drugs"][re.sub(r"\s+", " ", d_).strip()] += 1
        g["owned"] |= owned
        if not owned:
            g["partners"].add(lead.get("name"))
        kept += 1
    out, stale = [], []
    meta = {n: (a, rc, pk) for n, k, a, rc, pk in B}
    ACTIVE = {"RECRUITING", "ACTIVE_NOT_RECRUITING", "NOT_YET_RECRUITING", "ENROLLING_BY_INVITATION"}
    for b, g in prog.items():
        # keep programs still in development: an active trial, or one completed / unknown since 2025
        is_stale = not any(x["status"] in ACTIVE or x["pcd"][:4] >= "2025" for x in g["trials"])
        a, rc, pk = meta[b]
        stg = max(x["stage"] for x in g["trials"])
        lead_trials = sorted(g["trials"], key=lambda x: (-x["stage"], x["pcd"]))
        future = sorted(x["pcd"] for x in g["trials"] if x["pcd"] >= TODAY.isoformat()[:7] and x["status"] != "COMPLETED")
        pos = loa(stg, a, rc)
        txt = " ".join(x["title"] for x in g["trials"]) + " " + " ".join(g["conds"])
        niche = a == "Oncology" and NICHE.search(txt) and not BROAD.search(txt)
        if niche:
            pk = round(pk * 0.4)
        share = OWNED_MARGIN if g["owned"] else ROYALTY
        (stale if is_stale else out).append({"bucket": b, "area": a, "rate_class": rc or "area", "stage": stg, "pos": round(pos, 4), "peak": pk,
                    "owned": g["owned"], "partners": sorted(x for x in g["partners"] if x), "share": share,
                    "launch": YEAR0 + LAUNCH_LAG[stg], "drugs": [d_ for d_, _ in g["drugs"].most_common(3)],
                    "n_trials": len(g["trials"]), "next_readout": future[0] if future else None,
                    "lead_ncts": [x["nct"] for x in lead_trials[:3]], "lead_title": lead_trials[0]["title"],
                    "niche": bool(niche)})
    # one drug in several cardiometabolic buckets is one product with label expansions: keep the biggest bucket
    CM = {"Lipids / cardiometabolic", "Type 2 diabetes", "Thrombosis / cardiovascular", "Heart failure", "Hypertension",
          "Obesity", "MASH / liver fibrosis", "Kidney disease"}
    out.sort(key=lambda x: (-x["stage"], -x["peak"]))
    seen, merged = {}, []
    for x in out:
        lead = re.sub(r"[^a-z0-9]", "", (x["drugs"] or [""])[0].lower())[:10]
        if x["bucket"] in CM and lead and lead in seen:
            seen[lead]["label_expansions"] = seen[lead].get("label_expansions", []) + [x["bucket"]]
            continue
        if x["bucket"] in CM and lead:
            seen[lead] = x
        merged.append(x)
    out = merged
    # rank by a quick risked value proxy and keep the top programs
    out.sort(key=lambda x: -(x["peak"] * x["pos"] * x["share"] / (1.12 ** (x["launch"] - YEAR0))))
    return {"t": t, "name": name, "query": raw["query"], "n_studies": len(raw["studies"]), "n_used": kept,
            "programs": out[:MAX_PROGRAMS], "dropped_programs": [x["bucket"] for x in out[MAX_PROGRAMS:]],
            "stale_programs": [x["bucket"] for x in stale], "stale_full": stale}


OVR = {k: v for k, v in json.load(open(f"{D}/overrides.json")).items() if not k.startswith("_")} if os.path.exists(f"{D}/overrides.json") else {}

if __name__ == "__main__":
    os.makedirs(f"{D}/programs", exist_ok=True)
    TK = {x["t"]: x for x in json.load(open(f"{D}/../basic/tickers.json"))}
    ts = sys.argv[1:] or [l.strip() for l in open(f"{D}/universe.txt") if l.strip()]
    stat = Counter()
    for t in ts:
        raw = json.load(open(f"{D}/raw/{t}.json"))
        for extra in EXTRA_RAW.get(t, []):          # subsidiaries that sponsor the trials
            p = f"{D}/raw/{t}__{extra}.json"
            if os.path.exists(p):
                raw["studies"] += json.load(open(p))["studies"]
        if t == "SBFM":
            raw["studies"] = []          # "Sunshine" matched unrelated sponsors
        r = parse(t, raw, TK[t]["name"])
        ov = OVR.get(t, {})
        if ov:
            have = {p["bucket"] for p in r["programs"]}
            for p in r.get("stale_full", []):          # a hand-checked filing / approval revives a "stale" program
                if p["bucket"] in ov.get("set", {}) and p["bucket"] not in have:
                    r["programs"].append(p)
                    r["stale_programs"].remove(p["bucket"])
            r["programs"] = [p for p in r["programs"] if p["bucket"] not in ov.get("drop", [])]
            for p in r["programs"]:
                for k, v in ov.get("set", {}).get(p["bucket"], {}).items():
                    p[k] = v
            for p in ov.get("add", []):
                q = dict(p)
                q.setdefault("partners", [])
                q.setdefault("n_trials", 0); q.setdefault("next_readout", None); q.setdefault("lead_ncts", []); q.setdefault("lead_title", "")
                r["programs"].append(q)
            for p in r["programs"]:
                p["pos"] = round(loa(p["stage"], p["area"], None if p.get("rate_class") in (None, "area") else p["rate_class"]), 4)
                p.setdefault("share", OWNED_MARGIN if p["owned"] else ROYALTY)
                if "share" not in ov.get("set", {}).get(p["bucket"], {}):
                    p["share"] = OWNED_MARGIN if p["owned"] else ROYALTY
            r["overrides"] = True
            r["programs"].sort(key=lambda x: -(x["peak"] * x["pos"] * x["share"] / (1.12 ** max(0, x["launch"] - YEAR0))))
        r.pop("stale_full", None)
        json.dump(r, open(f"{D}/programs/{t}.json", "w"), indent=1)
        stat["with programs" if r["programs"] else "none"] += 1
        if len(ts) <= 12:
            print(f"== {t} {r['name']}  studies {r['n_studies']} used {r['n_used']}")
            for p in r["programs"]:
                print(f"   {p['bucket']:32s} stage {p['stage']:<3} PoS {p['pos']:.0%} peak {p['peak']:>5} "
                      f"{'own' if p['owned'] else 'partner'} launch {p['launch']} next {p['next_readout']} {p['drugs'][:2]}")
    print(dict(stat))
