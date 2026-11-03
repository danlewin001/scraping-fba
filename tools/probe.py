"""Sonde : télécharge une page avec un vrai navigateur et cherche EAN / GTIN dans tout le code."""
import asyncio, json, re, sys
from playwright.async_api import async_playwright

URL = sys.argv[1]

async def main():
    async with async_playwright() as pw:
        b = await pw.chromium.launch(headless=True)
        ctx = await b.new_context(locale="fr-BE", user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"))
        page = await ctx.new_page()
        jsons = []
        async def on_resp(r):
            if "json" in (r.headers.get("content-type") or ""):
                try:
                    t = await r.text()
                    if re.search(r"gtin|ean|barcode", t, re.I):
                        jsons.append((r.url, t[:3000]))
                except Exception:
                    pass
        page.on("response", on_resp)
        resp = await page.goto(URL, wait_until="networkidle", timeout=60000)
        print("STATUS", resp.status if resp else None)
        html = await page.content()
        print("TAILLE HTML", len(html))
        print("TITRE", await page.title())
    if len(html) < 5000:
        print("HTML:", html.replace("\n", " "))
        for m in re.finditer(r'.{0,80}(gtin\d*|\bean\b|barcode|eancode|ean13).{0,80}', html, re.I):
            print("MOTCLE:", m.group(0).replace("\n", " "))
        for blk in re.findall(r'<script[^>]+ld\+json[^>]*>(.*?)</script>', html, re.S):
            print("JSONLD:", blk.strip()[:1500])
        nums = sorted(set(re.findall(r'(?<!\d)(\d{13})(?!\d)', html)))
        print("NOMBRES 13 CHIFFRES:", nums[:30])
        for u, t in jsons[:5]:
            print("API JSON:", u, "\n", t[:1500])
        await b.close()

asyncio.run(main())
