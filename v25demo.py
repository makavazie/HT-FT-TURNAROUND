import time
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, simpledialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# ================== V25 HT/FT TURNAROUND SİSTEMİ — DEMO ==================
# Yeni Metodoloji: %60 Last6 + %40 Sezon Home/Away
# Paternler 59 maçlık eğitim verisiyle (06-13.03.2026) kalibre edildi
#
# V25 PATERN TABLOSU:
#   1/2-N  : Str≤-10, Asym≤15           → 9 TP / 0 YNL (eğitim seti)
#   2/1-N  : Str≥10, HT≥15, FT≥10      → 12 TP / 2 YNL — Prec ~%86 (eğitim seti)
#
# NOT: Bu demo sürümü — B şıkkı (daha fazla veri) ile kalibre edilecek
# ================== V21 HT/FT TURNAROUND SİSTEMİ ==================
# 166 maçlık arşiv verisine göre matematiksel olarak optimize edildi
# 
# ARŞİV ANALİZ SONUÇLARI (135 skorlu maç):
#   Toplam 2/1: 8 maç | 1/2: 5 maç | Olmadı: 122 maç
#
#   Gerçek 2/1 ortalama değerleri:
#     HT_Diff: ort=-4.9  med=-5.6  | min=-50.0  max=33.3
#     FT_Diff: ort=-2.1  med=+8.4  | min=-38.9  max=27.8
#     Asym:    ort=9.7   med=8.4   | min=0.0    max=22.2
#     Str:     ort=8.8   med=6.0   | min=-12    max=36
#
#   Gerçek 1/2 ortalama değerleri:
#     HT_Diff: ort=-10.0 med=-5.6  | min=-44.4  max=11.1
#     FT_Diff: ort=-25.5 med=-33.3 | min=-44.4  max=0.0
#     Asym:    ort=26.7  med=27.8  | min=11.1   max=50.0
#     Str:     ort=-8.8  med=0.0   | min=-38    max=10
#
# V20'nin SORUNLARI:
#   - 2/1 False Positive: 39/42 (sadece %7 başarı)
#   - 1/2 False Positive: 27/32 (sadece %16 başarı)
#   - 5 adet gerçek 2/1 PAS geçti (Asym çok düşüktü, eski sistem yakalayamadı)
#
# V21.1 DÜZELTMESİ:
#   STR_DOMINANT kuralındaki HT üst sınırı kaldırıldı.
#   Ingolstadt (HT=+33.3, Str=32) artık yakalanıyor.
#   Yeni kural: Str>=25 VE FT>=10 (HT değeri ne olursa olsun)
#   Prec: ~%14 (2 TP, 12 FP — 14 toplam sinyal)

CHROMEDRIVER_PATH = r"C:\Users\Makavazie\Desktop\pp\nowgoal\yeni\yeni\chromedriver.exe"
CHROME_BINARY = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
ARCHIVE_XLSX = "v25demo.xlsx"

# ── BİRLEŞİK ARŞİV DOSYALARI ──────────────────────────────────────────────
# Tüm geçmiş Excel arşivleri — komşu analizinde kullanılır
_NEIGHBOR_FILES = [
    # Önce proje içindeki dosyalar, sonra eski mutlak path'ler
    "mantikli.xlsx",
    "eğit.xlsx",
    "v25demo.xlsx",
    r"C:\Users\Makavazie\Desktop\pp\nowgoal\yeni\yeni\mantikli.xlsx",
    r"C:\Users\Makavazie\Desktop\pp\nowgoal\yeni\yeni\eğit.xlsx",
    r"C:\Users\Makavazie\Desktop\pp\nowgoal\yeni\yeni\v25demo.xlsx",
]

_neighbor_cache = None  # Performans için tek seferinde yükle
_model_cache = None     # Backtest ile kalibre edilen parametre cache'i

def _load_combined_archive():
    """Tüm arşiv dosyalarını yükle, birleştir ve cache'le."""
    global _neighbor_cache
    if _neighbor_cache is not None:
        return _neighbor_cache

    def flt(s):
        try: return float(str(s).replace('%','').replace('+','').strip())
        except: return None

    def parse_score(s):
        m = re.match(r'(\d+)[^0-9](\d+)', str(s).strip())
        return (int(m.group(1)), int(m.group(2))) if m else (None, None)

    def result_type(row):
        h1, a1 = parse_score(row.get('HT Score', ''))
        h2, a2 = parse_score(row.get('MS Score', ''))
        if None in (h1, a1, h2, a2): return 'unknown'
        if h1 > a1 and h2 < a2: return '1/2'
        if h1 < a1 and h2 > a2: return '2/1'
        return 'other'

    frames = []
    for path in _NEIGHBOR_FILES:
        try:
            df = pd.read_excel(path, engine='openpyxl', dtype=str)
            if 'FINAL' in df.columns:
                df.rename(columns={'FINAL': 'Final'}, inplace=True)
            df['actual'] = df.apply(result_type, axis=1)
            for col, key in [('HT Diff', 'HT Diff_f'), ('FT Diff', 'FT Diff_f'),
                              ('Asymmetry', 'Asymmetry_f'), ('Str Diff', 'Str Diff_f')]:
                df[key] = df[col].apply(flt)
            keep = ['Match', 'actual', 'HT Diff_f', 'FT Diff_f', 'Asymmetry_f', 'Str Diff_f']
            frames.append(df[[c for c in keep if c in df.columns]])
        except Exception:
            pass

    if not frames:
        _neighbor_cache = pd.DataFrame()
        return _neighbor_cache

    combined = pd.concat(frames, ignore_index=True)
    combined = combined[combined['actual'] != 'unknown'].drop_duplicates(subset=['Match'])
    _neighbor_cache = combined
    return _neighbor_cache


def find_neighbors(ht_diff, ft_diff, asym, str_diff, n=5):
    """
    Verilen profile en yakın N arşiv maçını Öklid mesafesiyle bul.
    Döner: [(sıra, sonuç_emoji, mesafe, Str, HT, FT, Asym, maç_adı), ...]
    """
    arc = _load_combined_archive()
    if arc.empty:
        return []

    query = {
        'HT Diff_f': ht_diff,
        'FT Diff_f': ft_diff,
        'Asymmetry_f': asym,
        'Str Diff_f': str_diff,
    }

    arc = arc.copy()
    arc['_dist'] = arc.apply(
        lambda r: sum((r[k] - v) ** 2 for k, v in query.items()
                      if pd.notna(r.get(k))) ** 0.5,
        axis=1
    )

    top = arc.nsmallest(n, '_dist')
    results = []
    for i, (_, row) in enumerate(top.iterrows(), 1):
        act = row['actual']
        if act == '1/2':   emoji = '✅1/2'
        elif act == '2/1': emoji = '✅2/1'
        else:              emoji = '🔵oth'
        results.append((
            i, emoji, row['_dist'],
            row.get('Str Diff_f', 0) or 0,
            row.get('HT Diff_f', 0) or 0,
            row.get('FT Diff_f', 0) or 0,
            row.get('Asymmetry_f', 0) or 0,
            row.get('Match', '?')
        ))
    return results


def _sigmoid(x):
    """Sayısal stabil, hızlı sigmoid."""
    if x > 50:
        return 1.0
    if x < -50:
        return 0.0
    import math
    return 1.0 / (1.0 + math.exp(-x))


def _estimate_probs_core(arc, ht_diff, ft_diff, asym, str_diff, params):
    """Verilen arşiv dataframe'i ile olasılık hesaplayıcı çekirdek."""
    feat_cols = ['HT Diff_f', 'FT Diff_f', 'Asymmetry_f', 'Str Diff_f']
    iqr_floor = float(params.get('iqr_floor', 10.0))
    k = int(params.get('k', 21))
    alpha = float(params.get('alpha', 0.65))  # kNN harman ağırlığı
    knn_other_weight = float(params.get('knn_other_weight', 0.60))

    # Robust ölçek (IQR) — outlier etkisini azaltır
    scales = {}
    for c in feat_cols:
        q1 = arc[c].quantile(0.25)
        q3 = arc[c].quantile(0.75)
        iqr = float(q3 - q1)
        scales[c] = iqr if iqr > 1e-6 else iqr_floor

    q = {'HT Diff_f': ht_diff, 'FT Diff_f': ft_diff, 'Asymmetry_f': asym, 'Str Diff_f': str_diff}
    arc2 = arc.copy()
    arc2['_dist'] = arc2.apply(
        lambda r: sum(((r[c] - q[c]) / scales[c]) ** 2 for c in feat_cols) ** 0.5,
        axis=1
    )

    top = arc2.nsmallest(max(5, min(k, len(arc2))), '_dist').copy()
    top['_w'] = 1.0 / (0.35 + top['_dist'])  # çok yakın komşular daha ağır
    tot_w = float(top['_w'].sum()) or 1.0

    w12 = float(top.loc[top['actual'] == '1/2', '_w'].sum())
    w21 = float(top.loc[top['actual'] == '2/1', '_w'].sum())
    wot = max(tot_w - (w12 + w21), 0.0)
    knn_12, knn_21, knn_other = w12 / tot_w, w21 / tot_w, wot / tot_w

    # Domain pattern score (soft rules): hard-threshold yerine puan yaklaşımı
    pat_12 = (
        0.35 * _sigmoid((-str_diff - 9) / 5.5) +
        0.25 * _sigmoid((10 - asym) / 4.5) +
        0.20 * _sigmoid((-ft_diff - 4) / 6.0) +
        0.20 * _sigmoid((-ht_diff + 2) / 7.0)
    )
    pat_21 = (
        0.30 * _sigmoid((str_diff - 9) / 5.0) +
        0.25 * _sigmoid((ht_diff - 4) / 6.0) +
        0.25 * _sigmoid((ft_diff - 10) / 6.0) +
        0.20 * _sigmoid((18 - asym) / 5.0)
    )

    p12 = alpha * knn_12 + (1 - alpha) * pat_12
    p21 = alpha * knn_21 + (1 - alpha) * pat_21

    # normalize
    s = p12 + p21 + (knn_other_weight * knn_other)
    if s <= 1e-9:
        return {
            'p_12': 0.0, 'p_21': 0.0, 'p_other': 1.0,
            'knn_12': knn_12, 'knn_21': knn_21, 'pat_12': pat_12, 'pat_21': pat_21
        }
    p12n = p12 / s
    p21n = p21 / s
    pother = max(0.0, 1.0 - (p12n + p21n))
    return {
        'p_12': p12n, 'p_21': p21n, 'p_other': pother,
        'knn_12': knn_12, 'knn_21': knn_21, 'pat_12': pat_12, 'pat_21': pat_21
    }


def _evaluate_backtest(arc, params):
    """
    Leave-one-out benzeri basit backtest skoru.
    Döner: (objective, metrics_dict)
    """
    min_train = 18
    tp = fp = fn = 0
    sig = 0
    turnaround_total = 0
    used = 0

    gate_p = float(params['gate_p'])
    gate_margin = float(params['gate_margin'])
    gate_other = float(params['gate_other'])

    for idx in range(len(arc)):
        row = arc.iloc[idx]
        actual = row['actual']
        is_turn = actual in ('1/2', '2/1')
        if is_turn:
            turnaround_total += 1

        train = arc.drop(arc.index[idx])
        if len(train) < min_train:
            continue

        probs = _estimate_probs_core(
            train,
            row['HT Diff_f'],
            row['FT Diff_f'],
            row['Asymmetry_f'],
            row['Str Diff_f'],
            params
        )
        p12, p21, pother = probs['p_12'], probs['p_21'], probs['p_other']
        top_dir = '1/2' if p12 >= p21 else '2/1'
        top_p = max(p12, p21)
        margin = abs(p12 - p21)
        signal = (top_p >= gate_p and margin >= gate_margin and pother <= gate_other)
        used += 1

        if signal:
            sig += 1
            if actual == top_dir:
                tp += 1
            else:
                fp += 1
        else:
            if is_turn:
                fn += 1

    precision = (tp / sig) if sig else 0.0
    recall = (tp / turnaround_total) if turnaround_total else 0.0
    signal_rate = (sig / used) if used else 0.0

    # Amaç: yüksek precision + makul recall + aşırı sinyal cezası
    objective = (
        precision * 0.62 +
        recall * 0.33 -
        max(signal_rate - 0.42, 0) * 0.10
    )
    metrics = {
        'tp': tp, 'fp': fp, 'fn': fn, 'signals': sig,
        'precision': precision, 'recall': recall, 'signal_rate': signal_rate, 'used': used
    }
    return objective, metrics


def calibrate_model_params():
    """
    eğit.xlsx + v25demo.xlsx (ve varsa mantikli.xlsx) üzerinden basit grid-search kalibrasyonu.
    Sonucu cache'ler.
    """
    global _model_cache
    if _model_cache is not None:
        return _model_cache

    arc = _load_combined_archive().copy()
    req = ['HT Diff_f', 'FT Diff_f', 'Asymmetry_f', 'Str Diff_f', 'actual']
    if arc.empty or any(c not in arc.columns for c in req):
        _model_cache = {
            'params': {'k': 21, 'alpha': 0.65, 'knn_other_weight': 0.60,
                       'gate_p': 0.47, 'gate_margin': 0.12, 'gate_other': 0.55,
                       'iqr_floor': 10.0},
            'metrics': {'precision': 0.0, 'recall': 0.0, 'signals': 0, 'used': 0},
            'calibrated': False
        }
        return _model_cache

    arc = arc[req].dropna()
    if len(arc) < 25:
        _model_cache = {
            'params': {'k': 21, 'alpha': 0.65, 'knn_other_weight': 0.60,
                       'gate_p': 0.47, 'gate_margin': 0.12, 'gate_other': 0.55,
                       'iqr_floor': 10.0},
            'metrics': {'precision': 0.0, 'recall': 0.0, 'signals': 0, 'used': len(arc)},
            'calibrated': False
        }
        return _model_cache

    best = None
    grid_k = [13, 17, 21, 25]
    grid_alpha = [0.55, 0.65, 0.75]
    grid_gate_p = [0.45, 0.47, 0.50, 0.53]
    grid_margin = [0.10, 0.12, 0.15]
    grid_gate_other = [0.50, 0.55, 0.60]

    for k in grid_k:
        for alpha in grid_alpha:
            for gp in grid_gate_p:
                for gm in grid_margin:
                    for go in grid_gate_other:
                        params = {
                            'k': k,
                            'alpha': alpha,
                            'knn_other_weight': 0.60,
                            'gate_p': gp,
                            'gate_margin': gm,
                            'gate_other': go,
                            'iqr_floor': 10.0
                        }
                        obj, met = _evaluate_backtest(arc, params)
                        # Çok az sinyal üretip şişirme olmasın
                        if met['signals'] < 6:
                            continue
                        cand = (obj, met['precision'], met['recall'], -met['fp'], params, met)
                        if best is None or cand > best:
                            best = cand

    if best is None:
        chosen = {'k': 21, 'alpha': 0.65, 'knn_other_weight': 0.60,
                  'gate_p': 0.47, 'gate_margin': 0.12, 'gate_other': 0.55,
                  'iqr_floor': 10.0}
        _, met = _evaluate_backtest(arc, chosen)
        _model_cache = {'params': chosen, 'metrics': met, 'calibrated': False}
        return _model_cache

    _model_cache = {
        'params': best[4],
        'metrics': best[5],
        'calibrated': True
    }
    return _model_cache


def estimate_turnaround_probabilities(ht_diff, ft_diff, asym, str_diff):
    """
    Arşiv + pre-match metriklerle 1/2 ve 2/1 olasılıklarını tahmin eder.
    Yöntem:
      1) Robust ölçekli distance-weighted kNN olasılığı
      2) Domain pattern skoru (form + güç tutarlılığı)
      3) Harmanlama (kNN %65, pattern %35)
    Döner:
      {p_12, p_21, p_other, knn_12, knn_21, pat_12, pat_21}
    """
    arc = _load_combined_archive().copy()
    req = ['HT Diff_f', 'FT Diff_f', 'Asymmetry_f', 'Str Diff_f', 'actual']
    if arc.empty or any(c not in arc.columns for c in req):
        return {
            'p_12': 0.0, 'p_21': 0.0, 'p_other': 1.0,
            'knn_12': 0.0, 'knn_21': 0.0, 'pat_12': 0.0, 'pat_21': 0.0
        }

    arc = arc[req].dropna()
    if len(arc) < 8:
        return {
            'p_12': 0.0, 'p_21': 0.0, 'p_other': 1.0,
            'knn_12': 0.0, 'knn_21': 0.0, 'pat_12': 0.0, 'pat_21': 0.0
        }

    model = calibrate_model_params()
    return _estimate_probs_core(arc, ht_diff, ft_diff, asym, str_diff, model['params'])


def format_neighbors(neighbors, decision):
    """Komşu tablosunu metin olarak formatla."""
    if not neighbors:
        return "  (Arşiv yüklenemedi)\n"

    # Sinyal yönüne göre TP sayısını hesapla
    if '2/1' in decision:
        tp_marker = '✅2/1'
    elif '1/2' in decision:
        tp_marker = '✅1/2'
    else:
        tp_marker = None

    tp_in_top3 = sum(1 for n in neighbors[:3] if n[1] == tp_marker) if tp_marker else 0
    tp_in_top5 = sum(1 for n in neighbors if n[1] == tp_marker) if tp_marker else 0

    lines = []
    lines.append(f"{'─'*80}")
    lines.append(f"🔍 EN YAKIN ARŞİV MAÇLARI (Öklid mesafesi — 4 boyut)")
    lines.append(f"{'─'*80}")
    lines.append(f"  {'#':<3} {'Sonuç':<8} {'Dist':>6}  {'Str':>5} {'HT':>8} {'FT':>8} {'Asym':>7}  Maç")
    lines.append(f"  {'─'*75}")
    for rank, emoji, dist, str_d, ht, ft, asym_v, match in neighbors:
        lines.append(
            f"  #{rank:<2}  {emoji:<8}  d={dist:>5.1f}"
            f"  Str={str_d:>+4.0f} HT={ht:>+6.1f}% FT={ft:>+6.1f}% Asym={asym_v:>5.1f}%"
            f"  {match}"
        )

    if tp_marker:
        lines.append(f"  {'─'*75}")
        if tp_in_top3 >= 2:
            verdict = f"💚 Güçlü  — İlk 3'te {tp_in_top3} TP"
        elif tp_in_top3 == 1:
            verdict = f"🟡 Orta   — İlk 3'te {tp_in_top3} TP  |  İlk 5'te {tp_in_top5} TP"
        elif tp_in_top5 >= 2:
            verdict = f"🟡 Zayıf  — İlk 3'te 0 TP  |  İlk 5'te {tp_in_top5} TP"
        else:
            verdict = f"🔴 Geç    — İlk 5'te {tp_in_top5} TP"
        lines.append(f"  Komşu Değerlendirmesi: {verdict}")

    return "\n".join(lines)

# ══════════════════════════════════════════════
#  V23 MATEMATİKSEL EŞİKLER — GERÇEK FİNAL
#  432 maçlık arşiv (v18main + v158main + v188main)
# ══════════════════════════════════════════════
#
#  DEĞİŞİKLİKLER (V22→V23):
#  ✂️ P1 KALDIRILDI     (N=26, Prec=%3.8 — gürültü)
#  ✂️ ÇAKIŞMA KALDIRILDI (N=20, Prec=%5.0 — gürültü)
#  🔧 P2 ISLAH: Str≥25+FT≥10+HT≤15 (Man City tipi sahte sinyaller engellendi %16.7)
#  ➕ P3 YENİ: HT≥30+Str≥30+FT≤40 → %50! (Ingolstadt, St.Gallen tipi)
#  🔧 P4 ISLAH: Str≥-10 alt sınır eklendi (Kfar Saba hatası önlendi)
#  🔧 P5 ISLAH: HT≥-10 şartı eklendi (4 yanlış sinyal önlendi, Prec→%25)
#  🔧 P6 ISLAH: HT≥-10 şartı eklendi (%10.5→%16.7, 2 yanlış önlendi)
#  ➕ N2 YENİ: HT≤-10+FT≤-20+Str≥10 → %16.7 (Marseille, Waldhof tipi)
#  - SIRALAMA: P9>P8>P6>P5>N2>P4>P3>P2
#
#  V22 ESKİ : ~130 sinyal → 14 TP → 2 YANLIŞ → %10.8 → 14/41=%34.1
#  V23 YENİ :   89 sinyal → 16 TP → 0 YANLIŞ → %18.0 → 16/41=%39.0
#  P9 ISLAH : Str≥-20 eklendi (Brighton Str=-34 önlendi, %50→%67)

# GÜVENİLİRLİK ORANLARI (432 maçlık arşiv — gerçek ölçüm)
RELIABILITY = {
    "1/2_A": {"rate": "~%22 (9/41)", "desc": "Str≤-15 + Asym≤15 — 0 YNL ⭐"},
    "1/2_B": {"rate": "~%22 (9/41)", "desc": "Str≤-15 + Asym≤15 — 0 YNL ⭐"},
    "2/1_Z": {"rate": "~%18 (10/56)", "desc": "Str≥10 + HT≥5 + FT 15~40 + Asym≤20 — 0 YNL ⭐"},
    "2/1_X1":{"rate": "~%18 (10/56)", "desc": "Str≥10 + HT≥5 + FT 15~40 + Asym≤20 — 0 YNL ⭐"},
    "2/1_X2":{"rate": "~%18 (10/56)", "desc": "Str≥10 + HT≥5 + FT 15~40 + Asym≤20 — 0 YNL ⭐"},
    "PAS":   {"rate": "—%",   "desc": "Sinyal yok"},
}

def get_driver():
    options = Options()
    options.binary_location = CHROME_BINARY
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    service = Service(CHROMEDRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)

def txt(soup, selector):
    el = soup.select_one(selector)
    return el.get_text(" ", strip=True) if el else ""

def parse_strength(soup):
    text = txt(soup, "#strengthChart")
    m = re.search(r"(\d{1,3})\s+(\d{1,3})", text)
    return (int(m.group(1)), int(m.group(2))) if m else (50, 50)

def parse_form_from_standings(soup):
    """
    Her takım için FT ve HT win oranını ağırlıklı ortalama ile hesapla:
      %60 × Last6 (pts/18)  +  %40 × Sezon (Home/Away pts/max)
    Ev takımı için Home satırı, dep takımı için Away satırı kullanılır.
    Döner: home_ft_pct, away_ft_pct, home_ht_pct, away_ht_pct  (0-100 arası float)
    """
    home_standings = soup.select_one('.team-table-home')
    away_standings = soup.select_one('.team-table-guest')
    if not home_standings or not away_standings:
        # Fallback: sabit %50
        return 50.0, 50.0, 50.0, 50.0

    W_L6 = 0.60   # Last 6 ağırlığı
    W_SN  = 0.40   # Sezon ağırlığı

    def _pts_rate(wins, draws, n):
        """pts/max_pts oranı (0.0–1.0)"""
        if n <= 0:
            return 0.5
        pts     = wins * 3 + draws
        max_pts = n * 3
        return pts / max_pts

    def extract_form(standings_table, side):
        """
        side = 'home'  → ev takımı için Home satırını sezon verisi olarak al
        side = 'away'  → dep takımı için Away satırını sezon verisi olarak al

        Döner: (ft_pct, ht_pct)  — 0-100 arası
        """
        # Varsayılan değerler (veri okunamazsa)
        l6_ft_pts  = 9;  l6_ht_pts  = 9   # Last 6 pts (max 18)
        sn_ft_rate = 0.5; sn_ht_rate = 0.5  # Sezon oranı

        row_label = 'Home' if side == 'home' else 'Away'

        try:
            rows = standings_table.find_all('tr')
            in_ht = False

            for row in rows:
                rtext = row.get_text()

                # HT bölümüne geçiş
                if 'HT' in rtext and 'Matches' in rtext:
                    in_ht = True
                    continue

                cols = [c.get_text(strip=True) for c in row.find_all('td')]
                if len(cols) < 5:
                    continue

                label = cols[0]

                # ── Last 6 satırı ──────────────────────────────
                if label == 'Last 6':
                    try:
                        pts = int(cols[7]) if len(cols) >= 8 else (
                            int(cols[2]) * 3 + int(cols[3])
                        )
                        pts = min(pts, 18)
                    except Exception:
                        try:
                            w = int(cols[2]); d = int(cols[3])
                            pts = min(w * 3 + d, 18)
                        except Exception:
                            pts = 9
                    if in_ht:
                        l6_ht_pts = pts
                    else:
                        l6_ft_pts = pts

                # ── Sezon (Home veya Away) satırı ──────────────
                elif label == row_label:
                    try:
                        n  = int(cols[1])
                        w  = int(cols[2])
                        d  = int(cols[3])
                        rate = _pts_rate(w, d, n)
                    except Exception:
                        rate = 0.5
                    if in_ht:
                        sn_ht_rate = rate
                    else:
                        sn_ft_rate = rate

        except Exception:
            pass

        # Ağırlıklı ortalama → 0-100 arası yüzde
        ft_pct = (l6_ft_pts / 18 * W_L6 + sn_ft_rate * W_SN) * 100
        ht_pct = (l6_ht_pts / 18 * W_L6 + sn_ht_rate * W_SN) * 100
        return ft_pct, ht_pct

    home_ft_pct, home_ht_pct = extract_form(home_standings, 'home')
    away_ft_pct, away_ht_pct = extract_form(away_standings, 'away')
    return home_ft_pct, away_ft_pct, home_ht_pct, away_ht_pct


def analyze_match(soup, url):
    home = txt(soup, "#fbheader .home")
    away = txt(soup, "#fbheader .guest")
    home_str, away_str = parse_strength(soup)

    # parse_form_from_standings artık doğrudan yüzde döndürüyor (0-100)
    home_ft_pct, away_ft_pct, home_ht_pct, away_ht_pct = parse_form_from_standings(soup)

    diff_ft_pct = home_ft_pct - away_ft_pct
    diff_ht_pct = home_ht_pct - away_ht_pct

    diff_str = home_str - away_str
    asymmetry = abs(diff_ht_pct - diff_ft_pct)

    decision = "PAS"
    level = "⚪ OYNAMA"
    confidence = ""
    pattern_key = "PAS"
    pattern_desc = ""

    # ════════════════════════════════════════════════════════
    #  V25 DEMO — YENİ METODOLOJİ PATERN SİSTEMİ
    #  Eğitim: 59 maç (06-13.03.2026) | %60 Last6 + %40 Sezon Home/Away
    #
    #  PATERNLER:
    #    1/2-N : Str≤-10 + Asym≤15  → Away güç+tutarlılık
    #    2/1-N : Str≥10 + HT≥15 + FT≥10 → Ev her metrikte üstün
    # ════════════════════════════════════════════════════════

    model_info = calibrate_model_params()
    model_params = model_info['params']
    model_metrics = model_info['metrics']
    probs = estimate_turnaround_probabilities(diff_ht_pct, diff_ft_pct, asymmetry, diff_str)
    p12 = probs['p_12']; p21 = probs['p_21']; pother = probs['p_other']
    knn12 = probs['knn_12']; knn21 = probs['knn_21']
    pat12 = probs['pat_12']; pat21 = probs['pat_21']

    top_dir = "1/2" if p12 >= p21 else "2/1"
    top_p = max(p12, p21)
    second_p = min(p12, p21)
    margin = top_p - second_p

    # Kalibrasyon: Yüksek diğer olasılığı veya düşük margin => PAS
    signal_gate = (
        top_p >= float(model_params['gate_p']) and
        margin >= float(model_params['gate_margin']) and
        pother <= float(model_params['gate_other'])
    )

    if signal_gate and top_dir == "1/2":
        decision = "İY/MS 1/2"; level = "⚠️ 1/2-PRO — PROBABILISTIC EDGE"
        confidence = "Yüksek" if top_p >= 0.60 else "Orta"
        pattern_key = "1/2_A"
        pattern_desc = (
            f"\n⚠️ İY/MS 1/2 — 1/2-PRO: Olasılık + Pattern Harmanı\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"  🎯 P(1/2)    : %{p12*100:.1f}\n"
            f"  🎯 P(2/1)    : %{p21*100:.1f}\n"
            f"  🎯 P(other)  : %{pother*100:.1f}\n"
            f"  📌 Margin    : %{margin*100:.1f}\n"
            f"  🔎 kNN(1/2)  : %{knn12*100:.1f}  |  Pattern(1/2): %{pat12*100:.1f}\n"
            f"  ℹ️ HT/FT/Asym/Str: {diff_ht_pct:+.1f}% / {diff_ft_pct:+.1f}% / {asymmetry:.1f}% / {diff_str:+d}\n"
            f"Mantık: Arşiv komşu yoğunluğu + yumuşak pattern puanı birlikte 1/2 lehine."
        )
    elif signal_gate and top_dir == "2/1":
        decision = "İY/MS 2/1"; level = "⚠️ 2/1-PRO — PROBABILISTIC EDGE"
        confidence = "Yüksek" if top_p >= 0.60 else "Orta"
        pattern_key = "2/1_X1"
        pattern_desc = (
            f"\n⚠️ İY/MS 2/1 — 2/1-PRO: Olasılık + Pattern Harmanı\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"  🎯 P(2/1)    : %{p21*100:.1f}\n"
            f"  🎯 P(1/2)    : %{p12*100:.1f}\n"
            f"  🎯 P(other)  : %{pother*100:.1f}\n"
            f"  📌 Margin    : %{margin*100:.1f}\n"
            f"  🔎 kNN(2/1)  : %{knn21*100:.1f}  |  Pattern(2/1): %{pat21*100:.1f}\n"
            f"  ℹ️ HT/FT/Asym/Str: {diff_ht_pct:+.1f}% / {diff_ft_pct:+.1f}% / {asymmetry:.1f}% / {diff_str:+d}\n"
            f"Mantık: Arşiv komşu yoğunluğu + yumuşak pattern puanı birlikte 2/1 lehine."
        )

    # PAS
    if decision == "PAS" and not pattern_desc:
        reliability_str = "—%"
        pattern_desc = (
            f"\n⚪ V24 — TURNAROUND BEKLENMİYOR\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"Eşik geçilmedi: Pmax=%{top_p*100:.1f}, margin=%{margin*100:.1f}, P(other)=%{pother*100:.1f}\n"
            f"  Model eşikleri: P≥{model_params['gate_p']:.2f}, M≥{model_params['gate_margin']:.2f}, O≤{model_params['gate_other']:.2f}\n"
            f"  HT={diff_ht_pct:+.1f}%  FT={diff_ft_pct:+.1f}%  Asym={asymmetry:.1f}%  Str={diff_str:+d}"
        )
    else:
        reliability_str = RELIABILITY[pattern_key]["rate"]

    # Komşu analizi
    neighbors = find_neighbors(diff_ht_pct, diff_ft_pct, asymmetry, diff_str, n=5)

    # ── Komşu Filtresi (k3 >= 1) ──────────────────────────────────────────
    # Sinyal varsa ilk 3 komşuda en az 1 TP zorunlu — yoksa PAS'a çevir
    if decision != "PAS":
        yön = "2/1" if "2/1" in decision else "1/2"
        emoji_tp = f"✅{yön}"
        k3_tp = sum(1 for n in neighbors[:3] if n[1] == emoji_tp)
        if k3_tp < 1:
            filtre_notu = (
                f"\n🚫 KOMŞU FİLTRESİ — Sinyal iptal edildi\n"
                f"━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                f"Patern tetiklendi ({pattern_key}) ama ilk 3 komşuda hiç {yön} TP bulunamadı (k3<1).\n"
                f"Arşiv desteği yetersiz → sinyal güvenilir değil.\n"
                f"  HT={diff_ht_pct:+.1f}%  FT={diff_ft_pct:+.1f}%  Asym={asymmetry:.1f}%  Str={diff_str:+d}"
            )
            decision      = "PAS"
            level         = "⚪ OYNAMA"
            confidence    = ""
            reliability_str = "—%"
            pattern_desc  = filtre_notu

    neighbor_text = format_neighbors(neighbors, decision)

    result_text = (
        f"{'═'*80}\n"
        f"🎯 V25 DEMO HT/FT TURNAROUND (Yeni Metodoloji — 59 maç eğitim)\n"
        f"{'═'*80}\n"
        f"📊 MAÇ    : {home} vs {away}\n"
        f"📋 HAM VERİLER:\n"
        f"   Strength  : {home_str} — {away_str}  (Fark: {diff_str:+d})\n"
        f"   FT Form   : {home_ft_pct:.1f}%  —  {away_ft_pct:.1f}%\n"
        f"               Fark: {diff_ft_pct:+.1f}%\n"
        f"   HT Form   : {home_ht_pct:.1f}%  —  {away_ht_pct:.1f}%\n"
        f"               Fark: {diff_ht_pct:+.1f}%\n"
        f"   Asimetri  : {asymmetry:.1f}%\n"
        f"   (Form = %60 Last6 + %40 Sezon Home/Away)\n"
        f"{'─'*80}\n"
        f"{pattern_desc}\n"
        f"{'─'*80}\n"
        f"📊 KARAR         : {decision}\n"
        f"📈 SEVİYE        : {level}\n"
        f"💎 GÜVENİLİRLİK  : {confidence}  |  Arşiv Başarı Oranı: {reliability_str}\n"
        f"{neighbor_text}\n"
        f"{'─'*80}\n"
        f"ℹ️  V25 DEMO — Yeni metodoloji | %60 Last6 + %40 Sezon Home/Away\n"
        f"    Motor: Backtest-kalibre Robust kNN + Soft Pattern + risk kapısı\n"
        f"    Parametreler: k={model_params['k']} | α={model_params['alpha']:.2f} | "
        f"P≥{model_params['gate_p']:.2f}, M≥{model_params['gate_margin']:.2f}, O≤{model_params['gate_other']:.2f}\n"
        f"    Backtest: prec=%{model_metrics.get('precision',0)*100:.1f} "
        f"| recall=%{model_metrics.get('recall',0)*100:.1f} | sinyal={model_metrics.get('signals',0)}\n"
        f"    Komşu filtresi: k3≥1 | ⚠️  DEMO — Veri birikince eşikler yeniden kalibre edilmeli\n"
        f"{'═'*80}"
    )

    save_to_archive(home, away, decision, level, confidence, pattern_key,
                    diff_ht_pct, diff_ft_pct, asymmetry, diff_str, url=url)
    return result_text


# ══════════════════════════════════════════════
#  EXCEL FORMAT SABITLERI
# ══════════════════════════════════════════════
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_COLS = ["Match", "Decision", "Level", "Confidence", "Pattern",
         "HT Diff", "FT Diff", "Asymmetry", "Str Diff", "HT Score", "MS Score", "Final", "URL"]

_COL_WIDTHS = {'A':38,'B':12,'C':26,'D':24,'E':22,
               'F':9,'G':9,'H':11,'I':9,'J':9,'K':9,'L':7,'M':55}

_HDR_BG      = "1F6B2E"
# Skor sonrası renkler
_GREEN       = "00C853"  # ✅ Sinyal doğru
_RED         = "FF1744"  # ❌ Sinyal yanlış
_YELLOW      = "FFD600"  # ⚠️  PAS ama turnaround oldu (kaçırdık)
_BLUE        = "1565C0"  # ℹ️  PAS ve normal bitti
_NO_SCORE    = "FFFFFF"  # Skor henüz yok
_WHITE       = "FFFFFF"
_BLACK       = "000000"
_THIN        = Side(style='thin', color='CCCCCC')
_BORDER      = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def _parse_score(s):
    import re
    m = re.match(r'(\d+)[^0-9](\d+)', str(s).strip())
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def _actual_result(ht_score, ms_score):
    """HT ve MS skorundan gerçek sonucu döndür: '1/2', '2/1', veya 'other'"""
    h1, a1 = _parse_score(ht_score)
    h2, a2 = _parse_score(ms_score)
    if None in (h1, a1, h2, a2):
        return None
    if h1 > a1 and h2 < a2:
        return '1/2'
    if h1 < a1 and h2 > a2:
        return '2/1'
    return 'other'


def _final_label(ht_score, ms_score):
    """1/1, 1/X, 1/2, X/1, X/X, X/2, 2/1, 2/X, 2/2"""
    h1, a1 = _parse_score(ht_score)
    h2, a2 = _parse_score(ms_score)
    if None in (h1, a1, h2, a2):
        return ""
    ht = "1" if h1 > a1 else ("X" if h1 == a1 else "2")
    ft = "1" if h2 > a2 else ("X" if h2 == a2 else "2")
    return f"{ht}/{ft}"


def _row_colors(decision, ht_score="", ms_score=""):
    d = str(decision)
    has_score = bool(str(ht_score).strip() and str(ms_score).strip()
                     and str(ht_score) not in ('nan','None','')
                     and str(ms_score) not in ('nan','None',''))

    if not has_score:
        # Skor yok — karar renginde göster (orijinal davranış)
        if '2/1' in d: return "FF4444", _WHITE
        if '1/2' in d: return "CC0000", _WHITE
        return _NO_SCORE, _BLACK

    actual = _actual_result(ht_score, ms_score)

    if '2/1' in d or '1/2' in d:
        # Sinyal verilmiş
        signal_dir = '2/1' if '2/1' in d else '1/2'
        if actual == signal_dir:
            return _GREEN, _BLACK   # ✅ Doğru
        else:
            return _RED, _WHITE     # ❌ Yanlış
    else:
        # PAS dendi
        if actual in ('1/2', '2/1'):
            return _YELLOW, _BLACK  # ⚠️  Kaçırdık
        else:
            return _BLUE, _WHITE    # ℹ️  PAS doğru


def _write_excel(df):
    """DataFrame'i formatlanmış Excel olarak yaz — her seferinde tam format."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(_COLS))}{len(df)+1}"

    # Başlık
    hdr_fill = PatternFill("solid", fgColor=_HDR_BG)
    hdr_font = Font(bold=True, color=_WHITE, name='Calibri', size=11)
    hdr_align = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22
    for ci, h in enumerate(_COLS, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = hdr_fill; c.font = hdr_font
        c.alignment = hdr_align; c.border = _BORDER

    # Veri satırları
    for ri, row in df.iterrows():
        er = ri + 2
        bg, fg = _row_colors(row.get('Decision', ''), row.get('HT Score', ''), row.get('MS Score', ''))
        fill = PatternFill("solid", fgColor=bg)
        ws.row_dimensions[er].height = 18
        row_data = dict(row)
        row_data['Final'] = _final_label(row.get('HT Score', ''), row.get('MS Score', ''))
        vals = [row_data.get(c, '') for c in _COLS]
        for ci, val in enumerate(vals, 1):
            c = ws.cell(row=er, column=ci)
            c.value = '' if str(val) in ('nan','None','NaN') else val
            c.fill = fill
            c.font = Font(bold=(ci <= 2), color=fg, name='Calibri', size=10)
            c.alignment = Alignment(
                horizontal='left' if ci == 1 else 'center',
                vertical='center')
            c.border = _BORDER

    for col_letter, width in _COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(ARCHIVE_XLSX)


def save_to_archive(home, away, decision, level, confidence, pattern_key,
                    ht_diff, ft_diff, asym, str_diff, url=""):
    try:
        df = pd.read_excel(ARCHIVE_XLSX, engine="openpyxl", dtype=str)
        for col in _COLS:
            if col not in df.columns:
                df[col] = ""
        df = df[_COLS]
    except FileNotFoundError:
        df = pd.DataFrame(columns=_COLS)

    new_row = {
        "Match":      f"{home} vs {away}",
        "Decision":   decision,
        "Level":      level,
        "Confidence": confidence,
        "Pattern":    pattern_key,
        "HT Diff":    f"{ht_diff:+.1f}%",
        "FT Diff":    f"{ft_diff:+.1f}%",
        "Asymmetry":  f"{asym:.1f}%",
        "Str Diff":   f"{str_diff:+d}",
        "HT Score":   "",
        "MS Score":   "",
        "URL":        url,
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    _write_excel(df)


def extract_match_id(url):
    """URL'den nowgoal maç ID'sini çıkar. Örn: .../2932464/... → 2932464"""
    m = re.search(r'[/-](\d{6,8})(?:[/-]|$)', url)
    return m.group(1) if m else None


def update_scores():
    try:
        df = pd.read_excel(ARCHIVE_XLSX, engine="openpyxl", dtype=str)
    except FileNotFoundError:
        messagebox.showerror("Hata", "Arşiv dosyası bulunamadı!")
        return

    for col in _COLS:
        if col not in df.columns:
            df[col] = ""
    df = df[_COLS]

    # Skoru eksik, URL'si olan maçları bul
    pending = df[
        (df["URL"].notna()) & (df["URL"] != "") &
        (df["HT Score"].isna() | (df["HT Score"] == "") |
         df["MS Score"].isna() | (df["MS Score"] == ""))
    ].copy()

    if len(pending) == 0:
        messagebox.showinfo("Bilgi", "Güncellenecek maç yok!\nURL'si olan tüm maçların skoru mevcut.")
        return

    # URL'lerden maç ID'lerini çıkar
    id_to_idx = {}
    for idx, row in pending.iterrows():
        mid = extract_match_id(str(row["URL"]))
        if mid:
            id_to_idx[mid] = idx

    if not id_to_idx:
        messagebox.showerror("Hata", "Hiçbir URL'den maç ID'si çıkarılamadı.")
        return

    output_box.delete(1.0, tk.END)
    output_box.insert(tk.END, f"⏳ {len(id_to_idx)} maç için nowgoal taranıyor...\n\n")
    root.update()

    NOWGOAL_HOME = "https://live5.nowgoal26.com/football/fixture?f=ft1"
    driver = get_driver()
    try:
        driver.get(NOWGOAL_HOME)
        time.sleep(6)
        soup = BeautifulSoup(driver.page_source, "html.parser")
    except Exception as e:
        driver.quit()
        messagebox.showerror("Hata", f"Ana sayfa açılamadı:\n{e}")
        return

    updated = 0
    not_found = []

    for mid, idx in id_to_idx.items():
        match_name = df.loc[idx, "Match"]
        row_el = soup.find(id=f"tr1_{mid}")
        if not row_el:
            not_found.append(match_name)
            output_box.insert(tk.END, f"  ⚠️  {match_name} — bulunamadı (ID={mid})\n")
            root.update()
            continue

        cells = row_el.find_all("td")
        try:
            # td[6]=MS (index 5), td[9]=HT (index 8)
            ms_raw = cells[5].get_text(strip=True)
            ht_raw = cells[8].get_text(strip=True)
            ms = re.sub(r"\s*-\s*", "-", ms_raw).strip()
            ht = re.sub(r"\s*-\s*", "-", ht_raw).strip()

            if re.match(r"^\d+-\d+$", ms) and re.match(r"^\d+-\d+$", ht):
                df.loc[idx, "HT Score"] = ht
                df.loc[idx, "MS Score"] = ms
                updated += 1
                output_box.insert(tk.END, f"  ✅ {match_name}  HT={ht}  MS={ms}\n")
            else:
                output_box.insert(tk.END, f"  ⚠️  {match_name} — format geçersiz ('{ms_raw}' / '{ht_raw}')\n")
                not_found.append(match_name)
        except IndexError:
            output_box.insert(tk.END, f"  ❌ {match_name} — hücre sayısı yetersiz\n")
            not_found.append(match_name)
        root.update()

    driver.quit()

    if updated > 0:
        _write_excel(df)

    output_box.insert(tk.END, f"\n{'='*50}\n")
    output_box.insert(tk.END, f"✅ Güncellendi : {updated} maç\n")
    if not_found:
        output_box.insert(tk.END, f"⚠️  Bulunamadı  : {len(not_found)} maç\n")
        output_box.insert(tk.END, "   (Henüz bitmemiş veya farklı günde olabilir)\n")
    root.update()


def show_stats():
    """Arşiv istatistiklerini göster"""
    try:
        df = pd.read_excel(ARCHIVE_XLSX, engine="openpyxl", dtype=str)
    except FileNotFoundError:
        messagebox.showinfo("Bilgi", "Henüz arşiv dosyası yok.")
        return

    total = len(df)
    scored = df[df["MS Score"].notna() & (df["MS Score"] != "")]

    def is_turnaround(row, direction):
        def ps(s):
            m = re.match(r'(\d+)[^0-9](\d+)', str(s).strip())
            return (int(m.group(1)), int(m.group(2))) if m else (None, None)
        h1, a1 = ps(row["HT Score"]) if pd.notna(row["HT Score"]) else (None, None)
        h2, a2 = ps(row["MS Score"]) if pd.notna(row["MS Score"]) else (None, None)
        if None in (h1, a1, h2, a2):
            return False
        if direction == "2/1":
            return h1 < a1 and h2 > a2
        if direction == "1/2":
            return h1 > a1 and h2 < a2
        return False

    stats_lines = [f"{'═'*50}", "📊 V21 ARŞİV İSTATİSTİKLERİ",
                   f"{'─'*50}", f"Toplam kayıt : {total}",
                   f"Skorlu maç   : {len(scored)}", f"{'─'*50}"]

    for dec in ["İY/MS 2/1", "İY/MS 1/2", "PAS"]:
        sub = scored[scored["Decision"] == dec]
        if dec == "PAS":
            stats_lines.append(f"PAS          : {len(sub)} maç")
            continue
        direction = dec.split()[-1]
        correct = sub[sub.apply(lambda r: is_turnaround(r, direction), axis=1)]
        pct = len(correct) / len(sub) * 100 if len(sub) > 0 else 0
        stats_lines.append(f"{dec:<14}: {len(correct)}/{len(sub)} doğru  ({pct:.1f}%)")

    stats_lines.append(f"{'═'*50}")
    messagebox.showinfo("Arşiv İstatistikleri", "\n".join(stats_lines))


def run_analysis():
    url = url_entry.get().strip()
    if not url:
        messagebox.showwarning("Uyarı", "Lütfen URL girin!")
        return
    output_box.delete(1.0, tk.END)
    output_box.insert(tk.END, "⏳ Analiz ediliyor...\n")
    root.update()
    try:
        driver = get_driver()
        driver.get(url)
        time.sleep(7)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        driver.quit()
        result = analyze_match(soup, url)
        output_box.delete(1.0, tk.END)
        output_box.insert(tk.END, result)
    except Exception as e:
        messagebox.showerror("Hata", f"Analiz sırasında hata:\n{str(e)}")


# ══════════════════════════════════════════════
#  ELİT GUI — TERMINAL / TRADING DESK ESTETİĞİ
# ══════════════════════════════════════════════

# ── Renk paleti ──────────────────────────────
BG_ROOT     = "#05080f"   # Derin gece mavisi
BG_PANEL    = "#080d18"   # Panel arka planı
BG_INPUT    = "#0c1220"   # Input alanı
BG_OUTPUT   = "#060b14"   # Terminal çıktı
BG_BTN_PRI  = "#0f172a"   # Ana buton
BG_BTN_SEC  = "#0c1220"   # İkincil buton
BORDER_DIM  = "#1a2540"   # Soluk kenarlık
BORDER_ACC  = "#c8a84b"   # Altın kenarlık (accent)
FG_TITLE    = "#c8a84b"   # Altın başlık
FG_SUB      = "#4a6080"   # Alt başlık
FG_LABEL    = "#6b8aaa"   # Etiket
FG_INPUT    = "#d4bc78"   # Input metni
FG_OUTPUT   = "#8fb8d4"   # Çıktı metni
FG_BTN_PRI  = "#c8a84b"   # Ana buton metni
FG_BTN_SEC  = "#4a6080"   # İkincil buton metni
FG_ACCENT   = "#c8a84b"   # Vurgu
FG_DIM      = "#1e2d45"   # Çok soluk
FG_STATUS   = "#2a4060"   # Status bar

FONT_TITLE   = ("Consolas", 20, "bold")
FONT_SUB     = ("Consolas", 9)
FONT_LABEL   = ("Consolas", 10, "bold")
FONT_INPUT   = ("Consolas", 11)
FONT_OUTPUT  = ("Consolas", 11)
FONT_BTN_PRI = ("Consolas", 11, "bold")
FONT_BTN_SEC = ("Consolas", 10)
FONT_STATUS  = ("Consolas", 8)

# ── Yardımcı: ince ayırıcı çizgi ─────────────
def separator(parent, color=BORDER_DIM, thickness=1):
    tk.Frame(parent, bg=color, height=thickness).pack(fill="x")

# ── Hover efekti ─────────────────────────────
def add_hover(btn, normal_bg, hover_bg, normal_fg, hover_fg):
    btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg, fg=hover_fg))
    btn.bind("<Leave>", lambda e: btn.configure(bg=normal_bg, fg=normal_fg))

# ── Kök pencere ──────────────────────────────
root = tk.Tk()
root.title("V24 · HT/FT TURNAROUND SYSTEM")
root.geometry("1640x1020")
root.minsize(1200, 700)
root.configure(bg=BG_ROOT)

# Pencereyi ortalama
root.update_idletasks()
sw = root.winfo_screenwidth()
sh = root.winfo_screenheight()
x  = (sw - 1200) // 2
y  = (sh - 700) // 2
root.geometry(f"1640x1020+{x}+{y}")

# ── ÜSTTE İNCE ALTIN ŞERİT ───────────────────
tk.Frame(root, bg=FG_ACCENT, height=2).pack(fill="x")

# ── HEADER ───────────────────────────────────
header = tk.Frame(root, bg=BG_PANEL, pady=0)
header.pack(fill="x")

header_inner = tk.Frame(header, bg=BG_PANEL)
header_inner.pack(fill="x", padx=36, pady=16)

# Sol: başlık grubu
left_hdr = tk.Frame(header_inner, bg=BG_PANEL)
left_hdr.pack(side=tk.LEFT)

tk.Label(left_hdr,
         text="V24  ·  HT/FT TURNAROUND",
         font=FONT_TITLE, bg=BG_PANEL, fg=FG_TITLE,
         anchor="w").pack(anchor="w")

tk.Label(left_hdr,
         text="648 maç · mantık tabanlı · komşu filtreli · %60 L6 + %40 sezon",
         font=FONT_SUB, bg=BG_PANEL, fg=FG_SUB,
         anchor="w").pack(anchor="w", pady=(3,0))

# Sağ: sistem metrikleri (küçük panel)
right_hdr = tk.Frame(header_inner, bg=BG_PANEL)
right_hdr.pack(side=tk.RIGHT)

metrics = [
    ("62",    "SİNYAL"),
    ("15",    "TP"),
    ("3",     "YANLIŞ"),
    ("%24.2", "PRECİSİON"),
]
for val, lbl in metrics:
    cell = tk.Frame(right_hdr, bg=BG_PANEL, padx=18)
    cell.pack(side=tk.LEFT)
    tk.Label(cell, text=val, font=("Consolas", 15, "bold"),
             bg=BG_PANEL, fg=FG_ACCENT).pack()
    tk.Label(cell, text=lbl, font=("Consolas", 7),
             bg=BG_PANEL, fg=FG_SUB).pack()

separator(header, color=BORDER_ACC, thickness=1)

# ── INPUT SATIRI ─────────────────────────────
input_row = tk.Frame(root, bg=BG_INPUT, pady=0)
input_row.pack(fill="x")

input_inner = tk.Frame(input_row, bg=BG_INPUT)
input_inner.pack(fill="x", padx=36, pady=14)

tk.Label(input_inner,
         text="URL", font=FONT_LABEL,
         bg=BG_INPUT, fg=FG_LABEL, width=5, anchor="w").pack(side=tk.LEFT)

# Altın bordürlü entry wrapper
entry_wrap = tk.Frame(input_inner, bg=BORDER_ACC, padx=1, pady=1)
entry_wrap.pack(side=tk.LEFT, fill="x", expand=True, padx=(8, 16))
url_entry = tk.Entry(entry_wrap, font=FONT_INPUT,
                     bg=BG_ROOT, fg=FG_INPUT,
                     insertbackground=FG_ACCENT,
                     relief="flat", bd=6,
                     highlightthickness=0)
url_entry.pack(fill="x")

# ── BUTON SATIRI ─────────────────────────────
btn_row = tk.Frame(root, bg=BG_ROOT)
btn_row.pack(fill="x", padx=36, pady=(10, 0))

def make_primary_btn(parent, text, cmd):
    btn = tk.Button(parent, text=text, command=cmd,
                    bg=FG_ACCENT, fg=BG_ROOT,
                    font=FONT_BTN_PRI,
                    padx=28, pady=10,
                    relief="flat", bd=0,
                    activebackground="#e8c96a",
                    activeforeground=BG_ROOT,
                    cursor="hand2")
    add_hover(btn, FG_ACCENT, "#e8c96a", BG_ROOT, BG_ROOT)
    return btn

def make_secondary_btn(parent, text, cmd):
    btn = tk.Button(parent, text=text, command=cmd,
                    bg=BG_BTN_SEC, fg=FG_LABEL,
                    font=FONT_BTN_SEC,
                    padx=22, pady=10,
                    relief="flat", bd=0,
                    activebackground=BORDER_DIM,
                    activeforeground=FG_ACCENT,
                    cursor="hand2",
                    highlightthickness=1,
                    highlightbackground=BORDER_DIM)
    add_hover(btn, BG_BTN_SEC, BORDER_DIM, FG_LABEL, FG_ACCENT)
    return btn

btn_analiz  = make_primary_btn(btn_row,   "▶  ANALİZ ET",        run_analysis)
btn_skor    = make_secondary_btn(btn_row,  "↻  SKOR GÜNCELLE",    update_scores)
btn_stats   = make_secondary_btn(btn_row,  "◈  İSTATİSTİKLER",    show_stats)

btn_analiz.pack(side=tk.LEFT, padx=(0, 10))
btn_skor.pack(side=tk.LEFT, padx=(0, 8))
btn_stats.pack(side=tk.LEFT)

# Enter tuşu ile analiz
url_entry.bind("<Return>", lambda e: run_analysis())

# ── İNCE AYIRICI ─────────────────────────────
tk.Frame(root, bg=BORDER_DIM, height=1).pack(fill="x", pady=(14, 0))

# ── ÇIKTI ALANI ──────────────────────────────
output_wrap = tk.Frame(root, bg=BG_ROOT)
output_wrap.pack(fill="both", expand=True, padx=36, pady=(0, 0))

output_box = scrolledtext.ScrolledText(
    output_wrap,
    font=FONT_OUTPUT,
    bg=BG_OUTPUT,
    fg=FG_OUTPUT,
    insertbackground=FG_ACCENT,
    selectbackground=BORDER_ACC,
    selectforeground=BG_ROOT,
    relief="flat", bd=0,
    padx=24, pady=20,
    spacing1=1, spacing3=1,
    wrap=tk.NONE,
)
output_box.pack(fill="both", expand=True)

# Yatay kaydırma çubuğu
h_scroll = tk.Scrollbar(output_wrap, orient=tk.HORIZONTAL,
                         command=output_box.xview,
                         bg=BG_PANEL, troughcolor=BG_ROOT,
                         activebackground=FG_ACCENT)
h_scroll.pack(fill="x")
output_box.configure(xscrollcommand=h_scroll.set)

# Renk etiketleri (çıktı içi vurgular)
output_box.tag_configure("gold",    foreground=FG_ACCENT,   font=("Consolas", 11, "bold"))
output_box.tag_configure("dim",     foreground=FG_DIM)
output_box.tag_configure("bright",  foreground="#d4e8f8")
output_box.tag_configure("green",   foreground="#4ade80")
output_box.tag_configure("red",     foreground="#f87171")

# ── STATUS BAR ───────────────────────────────
tk.Frame(root, bg=BORDER_DIM, height=1).pack(fill="x")
status_bar = tk.Frame(root, bg=BG_PANEL, pady=5)
status_bar.pack(fill="x")

status_inner = tk.Frame(status_bar, bg=BG_PANEL)
status_inner.pack(fill="x", padx=36)

tk.Label(status_inner,
         text="V24  ©  2026  —  648 maçlık arşiv  ·  mantık tabanlı patern sistemi  ·  komşu filtresi aktif",
         font=FONT_STATUS, bg=BG_PANEL, fg=FG_STATUS,
         anchor="w").pack(side=tk.LEFT)

# Sağda canlı saat
clock_lbl = tk.Label(status_inner, text="", font=FONT_STATUS,
                      bg=BG_PANEL, fg=FG_SUB)
clock_lbl.pack(side=tk.RIGHT)

def tick():
    import datetime
    clock_lbl.config(text=datetime.datetime.now().strftime("%H:%M:%S"))
    root.after(1000, tick)
tick()

# ── ALTTA İNCE ALTIN ŞERİT ───────────────────
tk.Frame(root, bg=FG_ACCENT, height=1).pack(fill="x", side=tk.BOTTOM)

root.mainloop()
