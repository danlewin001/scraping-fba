"""
Scraper hebdomadaire Carrefour.be — rayons Jouets, Soins & hygiène, Non-alimentaire.

Récupère pour chaque produit : nom, marque, EAN, prix, ancien prix, % de réduction,
disponibilité, URL. Compare avec la semaine précédente et envoie un Excel par mail.

Variables d'environnement (à mettre dans les Secrets GitHub) :
  PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS   -> IPRoyal (résidentiel, Belgique)
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, MAIL_TO -> envoi du rapport
Optionnel :
  MAX_PRODUCTS (test : limite le nombre de fiches), CONCURRENCY (défaut 8)
"""

import asyncio
import json
import os
import random
import re
import smtplib
import sys
from datetime import date
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import urljoin, urlparse, urlencode, parse_qsl, urlunparse

import pandas as pd
from playwright.async_api import async_playwright

BASE = "https://www.carrefour.be"

# Points de départ des rayons. Le script suit automatiquement les sous-rayons
# dont l'URL commence par un de ces chemins.
START_PATHS = [
    "/fr/non-alimentaire",            # inclut /fr/non-alimentaire/jouet-et-loisirs
    "/fr/non-alimentaire/jouet-et-loisirs",
    "/fr/soins-et-hygiene",
]
# Pagination du site : ?p=2, ?p=3…

PRODUCT_URL_RE = re.compile(r"^/fr/[^/?#]+/\d{6,10}\.html$")
PERCENT_RE = re.compile(r"-\s?(\d{1,2})\s?%")
PRICE_RE = re.compile(r"(\d{1,4}(?:[.,]\d{2}))\s?€|€\s?(\d{1,4}(?:[.,]\d{2}))")

# Offres multi-achat : « 1+1 gratuit », « 2+1 offert », « 3 pour 2 », « 2e à -50% », « 2ème gratuit »
MULTI_PATTERNS = [
    # (regex, fonction -> (nb articles achetés, nb articles payés équivalents))
    (re.compile(r"\b(\d)\s?\+\s?(\d)\s?(?:gratuits?|offerts?)", re.I),
     lambda a, b: (int(a) + int(b), int(a))),
    (re.compile(r"\b(\d)\s?(?:pour|=)\s?(?:le\s?prix\s?de\s?)?(\d)\b"
                r"(?!\s?(?:pers|enf|adult|ans?\b|mois|kg|g\b|gr|ml|cl|l\b|cm|mm|m\b|x\b|pcs|pi[eè]ce|places?|joueurs?))",
                re.I),
     lambda n, k: (int(n), int(k)) if int(n) > int(k) else None),
    (re.compile(r"\b(\d)\s?(?:e|è|ème|eme|de)\s?(?:à|a)\s?-\s?(\d{1,3})\s?%", re.I),
     lambda n, p: (int(n), int(n) - int(p) / 100)),
    (re.compile(r"\b(\d)\s?(?:e|è|ème|eme|de)\s?(?:gratuit|offert)", re.I),
     lambda n: (int(n), int(n) - 1)),
]


def detect_multibuy(text):
    """Retourne (texte de l'offre, nb à acheter, % de réduction effective) ou (None, None, None)."""
    best = (None, None, None)
    for rx, fn in MULTI_PATTERNS:
        for m in rx.finditer(text):
            res = fn(*m.groups())
            if not res:
                continue
            qty, paid = res
            if qty < 2 or paid >= qty:
                continue
            pct = round(100 * (1 - paid / qty))
            if best[2] is None or pct > best[2]:
                best = (m.group(0).strip(), qty, pct)
    return best

DATA_DIR = Path("data")
PREV_FILE = DATA_DIR / "previous.csv"
OUT_XLSX = Path(f"carrefour_{date.today().isoformat()}.xlsx")

CONCURRENCY = int(os.getenv("CONCURRENCY", "8"))
MAX_PRODUCTS = int(os.getenv("MAX_PRODUCTS", "0")) or None


# ---------------------------------------------------------------- utilitaires

def to_float(txt):
    if txt is None:
        return None
    if isinstance(txt, (int, float)):
        return float(txt)
    m = re.search(r"\d+(?:[.,]\d+)?", str(txt).replace("\xa0", " "))
    return float(m.group(0).replace(",", ".")) if m else None


def with_params(url, **params):
    parts = urlparse(url)
    q = dict(parse_qsl(parts.query))
    q.update({k: str(v) for k, v in params.items()})
    return urlunparse(parts._replace(query=urlencode(q)))


def in_scope(path):
    return any(path.startswith(p) for p in START_PATHS)


def proxy_config():
    host = os.getenv("PROXY_HOST")
    if not host:
        print("⚠️  Pas de proxy configuré : connexion directe (risque de blocage).")
        return None
    return {
        "server": f"http://{host}:{os.getenv('PROXY_PORT', '12321')}",
        "username": os.getenv("PROXY_USER", ""),
        "password": os.getenv("PROXY_PASS", ""),
    }


async def polite_pause():
    await asyncio.sleep(random.uniform(0.8, 2.0))


# ---------------------------------------------------------------- extraction

def parse_product_html(html, url):
    """Extrait les infos d'une fiche produit. Priorité au JSON-LD schema.org,
    puis repli sur des motifs dans le HTML."""
    row = {"url": url, "id": re.search(r"/(\d{6,10})\.html", url).group(1)}

    # 1) JSON-LD
    for block in re.findall(
        r'<script[^>]+application/ld\+json[^>]*>(.*?)</script>', html, re.S | re.I
    ):
        try:
            data = json.loads(block.strip())
        except json.JSONDecodeError:
            continue
        items = data if isinstance(data, list) else data.get("@graph", [data])
        for it in items:
            if not isinstance(it, dict) or "Product" not in str(it.get("@type")):
                continue
            row["nom"] = it.get("name")
            brand = it.get("brand")
            row["marque"] = brand.get("name") if isinstance(brand, dict) else brand
            row["ean"] = (it.get("gtin13") or it.get("gtin") or it.get("gtin14")
                          or it.get("gtin12") or it.get("ean"))
            offers = it.get("offers") or {}
            if isinstance(offers, list):
                offers = offers[0] if offers else {}
            row["prix"] = to_float(offers.get("price") or offers.get("lowPrice"))
            avail = str(offers.get("availability", ""))
            row["dispo"] = "oui" if "InStock" in avail else ("non" if avail else None)

    # 2) Replis HTML
    if not row.get("nom"):
        m = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.S)
        row["nom"] = re.sub(r"<[^>]+>|\s+", " ", m.group(1)).strip() if m else None
    if not row.get("ean"):
        m = re.search(r'(?:ean|gtin)[^0-9]{0,40}(\d{13})', html, re.I)
        row["ean"] = m.group(1) if m else None
    if row.get("prix") is None:
        m = re.search(r'itemprop="price"[^>]*content="([\d.,]+)"', html)
        row["prix"] = to_float(m.group(1)) if m else None

    # 3) Offres multi-achat (1+1 gratuit, 2e à -50 %, 3 pour 2…)
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html[:300000], flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&egrave;", "è", re.sub(r"&nbsp;|\s+", " ", text))
    row["offre_multi"], row["qte_offre"], row["reduction_effective_pct"] = detect_multibuy(text)

    # 4) Réduction directe : badge « -XX% » (hors offres multi-achat) et ancien prix barré
    for rx, _ in MULTI_PATTERNS:
        text = rx.sub(" ", text)
    pct = PERCENT_RE.findall(text)
    row["reduction_pct"] = max(int(p) for p in pct) if pct else None
    m = re.search(
        r'class="[^"]*(?:strike|old|was|standard|barr|crossed|list-price)[^"]*"[^>]*>(.*?)<',
        html, re.S | re.I)
    row["ancien_prix"] = to_float(PRICE_RE.search(m.group(1)).group(0)) if m and PRICE_RE.search(m.group(1)) else None
    if row["reduction_pct"] is None and row["ancien_prix"] and row.get("prix"):
        if row["ancien_prix"] > row["prix"]:
            row["reduction_pct"] = round(100 * (1 - row["prix"] / row["ancien_prix"]))
    if row["ancien_prix"] is None and row["reduction_pct"] and row.get("prix"):
        row["ancien_prix"] = round(row["prix"] / (1 - row["reduction_pct"] / 100), 2)
    finalize_effective(row)
    return row


def apply_tile(row, tile_text):
    """La tuile du rayon est la source la plus fiable pour l'offre de CE produit
    (la fiche affiche aussi des produits recommandés)."""
    offre, qte, eff = detect_multibuy(tile_text)
    if offre:
        row["offre_multi"], row["qte_offre"], row["reduction_effective_pct"] = offre, qte, eff
    clean = tile_text
    for rx, _ in MULTI_PATTERNS:
        clean = rx.sub(" ", clean)
    pct = PERCENT_RE.findall(clean)
    if pct:
        row["reduction_pct"] = max(int(p) for p in pct)
        if row.get("prix") and not row.get("ancien_prix"):
            row["ancien_prix"] = round(row["prix"] / (1 - row["reduction_pct"] / 100), 2)
    if offre or pct:
        if not offre and row.get("offre_multi") is None:
            row["reduction_effective_pct"] = None
        finalize_effective(row)


def finalize_effective(row):
    """Calcule le prix unitaire réel en achetant la quantité de l'offre."""
    eff = row.get("reduction_effective_pct")
    prix = row.get("prix")
    if eff and prix:
        row["prix_unitaire_effectif"] = round(prix * (1 - eff / 100), 2)
    else:
        row["reduction_effective_pct"] = row.get("reduction_pct")
        row["prix_unitaire_effectif"] = prix


TILE_JS = r"""
() => {
  const out = {};
  for (const a of document.querySelectorAll('a[href]')) {
    let href;
    try { href = new URL(a.href, location.origin).pathname; } catch { continue; }
    if (!/^\/fr\/[^/]+\/\d{6,10}\.html$/.test(href)) continue;
    // remonte jusqu'à la tuile produit pour lire le badge de réduction
    // texte de la tuile (badge -XX%, « 1+1 gratuit »…), sans déborder sur les tuiles voisines
    let tile = a, txt = '';
    for (let i = 0; i < 6 && tile.parentElement; i++) {
      const up = tile.parentElement;
      const n = up.querySelectorAll('a[href$=".html"]').length;
      if (n > 3) break;
      tile = up;
    }
    txt = (tile.innerText || '').replace(/\s+/g, ' ').slice(0, 400);
    if (!(href in out) || txt.length > (out[href] || '').length) out[href] = txt;
  }
  const cats = [...document.querySelectorAll('a[href]')]
    .map(a => { try { return new URL(a.href, location.origin).pathname; } catch { return null; } })
    .filter(Boolean);
  return {products: out, links: [...new Set(cats)]};
}
"""


# ---------------------------------------------------------------- crawl

async def crawl_listings(context):
    """Parcourt les pages rayons (+ pagination) et retourne
    {chemin_produit: % réduction vu sur la tuile}."""
    products, seen_cats = {}, set()
    queue = list(dict.fromkeys(START_PATHS))
    page = await context.new_page()
    missing = []

    while queue:
        path = queue.pop(0)
        if path in seen_cats:
            continue
        seen_cats.add(path)
        p = 1
        while True:
            url = BASE + path if p == 1 else with_params(BASE + path, p=p)
            try:
                resp = await page.goto(url, wait_until="domcontentloaded", timeout=45000)
                status = resp.status if resp else 0
            except Exception as e:
                print(f"  ✗ {url} : {e}")
                break
            if status >= 400:
                print(f"  ✗ {path} -> HTTP {status}")
                if p == 1:
                    missing.append(f"{path} (HTTP {status})")
                break
            await page.wait_for_timeout(1500)
            res = await page.evaluate(TILE_JS)
            new = {k: v for k, v in res["products"].items() if k not in products}
            for k, v in res["products"].items():
                if v or k not in products:
                    products[k] = v or products.get(k)
            for link in res["links"]:
                clean = link.split("?")[0].rstrip("/")
                if in_scope(clean) and not PRODUCT_URL_RE.match(clean) and clean not in seen_cats:
                    queue.append(clean)
            print(f"  {path} page {p}: {len(res['products'])} produits ({len(new)} nouveaux)")
            if not new:
                break
            p += 1
            if MAX_PRODUCTS and len(products) >= MAX_PRODUCTS:
                break
            await polite_pause()
        if MAX_PRODUCTS and len(products) >= MAX_PRODUCTS:
            break

    await page.close()
    return products, missing


async def fetch_products(context, paths, tile_pct):
    sem = asyncio.Semaphore(CONCURRENCY)
    rows, errors = [], 0

    async def one(path):
        nonlocal errors
        async with sem:
            page = await context.new_page()
            try:
                resp = await page.goto(BASE + path, wait_until="domcontentloaded", timeout=45000)
                if resp and resp.status >= 400:
                    raise RuntimeError(f"HTTP {resp.status}")
                html = await page.content()
                row = parse_product_html(html, BASE + path)
                apply_tile(row, tile_pct.get(path) or "")
                rows.append(row)
            except Exception as e:
                errors += 1
                print(f"  ✗ {path} : {e}")
            finally:
                await page.close()
            await polite_pause()

    await asyncio.gather(*(one(p) for p in paths))
    return rows, errors


# ---------------------------------------------------------------- rapport

def build_report(df):
    cols = ["ean", "nom", "marque", "prix", "ancien_prix", "reduction_pct",
            "offre_multi", "qte_offre", "reduction_effective_pct", "prix_unitaire_effectif",
            "dispo", "url", "id"]
    df = df.reindex(columns=cols)
    nouveaux = baisses = pd.DataFrame(columns=cols)

    if PREV_FILE.exists():
        prev = pd.read_csv(PREV_FILE, dtype={"id": str, "ean": str})
        if "prix_unitaire_effectif" not in prev:
            prev["prix_unitaire_effectif"] = prev["prix"]
        nouveaux = df[~df["id"].isin(prev["id"])]
        merged = df.merge(prev[["id", "prix_unitaire_effectif"]], on="id",
                          suffixes=("", "_semaine_passee"))
        merged["baisse_pct"] = (100 * (1 - merged["prix_unitaire_effectif"]
                                       / merged["prix_unitaire_effectif_semaine_passee"])).round(1)
        baisses = merged[merged["baisse_pct"] >= 5].sort_values("baisse_pct", ascending=False)

    promos = df[df["reduction_effective_pct"].fillna(0) > 0].sort_values(
        "reduction_effective_pct", ascending=False)
    multi = df[df["offre_multi"].notna()].sort_values("reduction_effective_pct", ascending=False)

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        promos.to_excel(xw, sheet_name="Promos", index=False)
        multi.to_excel(xw, sheet_name="Offres x+x", index=False)
        baisses.to_excel(xw, sheet_name="Baisses de prix", index=False)
        nouveaux.to_excel(xw, sheet_name="Nouveautés", index=False)
        df.to_excel(xw, sheet_name="Tous les produits", index=False)
        pd.DataFrame({"EAN": df["ean"].dropna().unique()}).to_excel(xw, sheet_name="EAN (pour SAS)", index=False)

    DATA_DIR.mkdir(exist_ok=True)
    df.to_csv(PREV_FILE, index=False)
    return len(promos), len(multi), len(baisses), len(nouveaux)


def send_mail(summary):
    host, to = os.getenv("SMTP_HOST"), os.getenv("MAIL_TO")
    if not (host and to):
        print("Pas de config mail : rapport seulement en artefact local.")
        return
    msg = EmailMessage()
    msg["Subject"] = f"Carrefour.be — scraping du {date.today():%d/%m/%Y}"
    msg["From"] = os.getenv("SMTP_USER")
    msg["To"] = to
    msg.set_content(summary)
    msg.add_attachment(OUT_XLSX.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=OUT_XLSX.name)
    with smtplib.SMTP(host, int(os.getenv("SMTP_PORT", "587"))) as s:
        s.starttls()
        s.login(os.getenv("SMTP_USER"), os.getenv("SMTP_PASS"))
        s.send_message(msg)
    print(f"📧 Rapport envoyé à {to}")


# ---------------------------------------------------------------- main

async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True, proxy=proxy_config())
        context = await browser.new_context(
            locale="fr-BE",
            timezone_id="Europe/Brussels",
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"),
            viewport={"width": 1366, "height": 900},
        )
        # Ne charge ni images, ni polices, ni vidéos : moins de Go consommés chez IPRoyal
        await context.route(
            re.compile(r".*\.(png|jpe?g|webp|gif|svg|woff2?|ttf|mp4|ico)(\?.*)?$"),
            lambda r: r.abort())

        print("1/3 Parcours des rayons…")
        tile_pct, missing = await crawl_listings(context)
        paths = list(tile_pct)[:MAX_PRODUCTS] if MAX_PRODUCTS else list(tile_pct)
        print(f"   → {len(paths)} produits trouvés")
        if not paths:
            await browser.close()
            sys.exit("Aucun produit trouvé : site bloqué ou structure changée.")

        print("2/3 Lecture des fiches produits…")
        rows, errors = await fetch_products(context, paths, tile_pct)
        await browser.close()

    print("3/3 Rapport…")
    df = pd.DataFrame(rows).drop_duplicates("id")
    n_promo, n_multi, n_baisse, n_new = build_report(df)
    summary = (
        f"Scraping Carrefour.be du {date.today():%d/%m/%Y}\n\n"
        f"Produits lus : {len(df)} (erreurs : {errors})\n"
        f"Avec EAN : {df['ean'].notna().sum()}\n"
        f"En promo : {n_promo} (dont offres x+x / 2e à -% : {n_multi})\n"
        f"Baisses de prix ≥ 5 % vs semaine passée : {n_baisse}\n"
        f"Nouveautés : {n_new}\n"
    )
    if missing:
        summary += "\n⚠️ Rayons introuvables (URL à corriger) :\n- " + "\n- ".join(missing) + "\n"
    print(summary)
    send_mail(summary)


if __name__ == "__main__":
    asyncio.run(main())
