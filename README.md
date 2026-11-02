# Scraping hebdo Carrefour.be

Chaque lundi matin, GitHub lance `scraper.py` qui parcourt les rayons
**Jouets**, **Soins & hygiène** et **Non-alimentaire** de Carrefour.be via un
proxy résidentiel belge IPRoyal, puis t'envoie un Excel par mail.

## Contenu de l'Excel
| Onglet | Contenu |
|---|---|
| Promos | toutes les réductions (-XX % et offres x+x), triées par réduction réelle |
| Offres x+x | « 1+1 gratuit », « 2+3 gratuit », « 3 pour 2 », « 2e à -50 % »… avec quantité à acheter et prix unitaire réel |
| Baisses de prix | baisses ≥ 5 % du prix unitaire réel par rapport à la semaine précédente |
| Nouveautés | produits apparus depuis la semaine précédente |
| Tous les produits | EAN, nom, marque, prix, ancien prix, % réduction, dispo, lien |
| EAN (pour SAS) | liste d'EAN prête à coller dans SellerAmp / Keepa |

## Installation (une seule fois)

### 1. Déposer les fichiers
- `scraper.py`, `requirements.txt`, `README.md` → **Add file → Upload files**
- `scrape.yml` → **Add file → Create new file**, nom :
  `.github/workflows/scrape.yml`, puis coller le contenu du fichier.

### 2. Ajouter les secrets
**Settings → Secrets and variables → Actions → New repository secret** :

| Nom | Valeur |
|---|---|
| `PROXY_HOST` | hôte IPRoyal (ex. `geo.iproyal.com`) |
| `PROXY_PORT` | port IPRoyal (ex. `12321`) |
| `PROXY_USER` | identifiant IPRoyal |
| `PROXY_PASS` | mot de passe IPRoyal **tel que généré avec le pays Belgique** |
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | ton adresse Gmail |
| `SMTP_PASS` | un **mot de passe d'application** Gmail (pas ton vrai mot de passe) |
| `MAIL_TO` | l'adresse qui reçoit le rapport |

### 3. Premier test
**Actions → Scraping Carrefour.be → Run workflow**, mettre `50` dans
« nombre max de produits ». Vérifier le mail reçu, puis relancer avec `0`
(tout) ou attendre lundi.

## Si ça bloque
- Mail « Rayons introuvables » → l'URL d'un rayon a changé (`START_PATHS` dans `scraper.py`).
- « Aucun produit trouvé » → Carrefour bloque : vérifier les secrets IPRoyal et le crédit restant.
- Peu de fiches ou beaucoup d'erreurs → baisser `CONCURRENCY` (défaut 8).
