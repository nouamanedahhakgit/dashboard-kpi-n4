# Dashboard KPI — Parc Conteneurs Navis N4

Outil de statistiques et de visualisation du parc conteneurs à partir de l'**API Universal Query de Navis N4** (authentification HTTP Basic).

Trois interfaces, un même moteur :

| Fichier | Interface | Usage |
|---|---|---|
| `kpi_web.py` + `index.html` | **Web** (navigateur, `localhost:8000`) | KPI + graphiques + filtres interactifs. **Recommandé.** |
| `kpi_gui.py` | Fenêtre Tkinter | Filtres + export Excel, sans navigateur |
| `kpi_stats.py` | Ligne de commande | Génère un classeur Excel (résumé + graphiques + données) |

## Prérequis

- Python 3.8+
- `pip install requests openpyxl`
  (Tkinter est inclus avec Python ; le serveur web n'utilise que la bibliothèque standard.)

## Configuration

```bash
cp .env.example .env
```

Puis renseigne dans `.env` :

```
N4_USER=ton_identifiant
N4_PASSWORD=ton_mot_de_passe
N4_URL=http://SERVEUR:PORT/apex/api/query?filtername=units_1&operatorId=...&complexId=...&facilityId=...&yardId=...
```

> `.env` n'est jamais versionné (voir `.gitignore`). Ne le partage pas.

## Lancer

### Interface web (recommandée)
```bash
python kpi_web.py
```
Le navigateur s'ouvre sur http://localhost:8000. Le serveur fait l'auth + l'appel API côté serveur (pas de souci CORS), met les données en cache (`cache.xml`), puis **le filtre pilote la requête** : le navigateur envoie les critères, le serveur renvoie uniquement les KPI et les agrégats des graphiques.

### Export Excel (CLI)
```bash
python kpi_stats.py                 # lit .env, appelle l'API, génère un .xlsx
python kpi_stats.py --in fichier.xml  # hors ligne, à partir d'un export XML
```

### Fenêtre Tkinter
```bash
python kpi_gui.py
```

## Indicateurs

Catégorie (import / export / storage), plein / vide, top armateurs, top POD,
reefers, dangereux (IMDG), hors gabarit, blocages navire / route, et
**temps de séjour (Dwell, en jours)** : médiane, p90, aging, histogramme.

Filtres : période (Last Move / EC-In Time / Complex InTime), catégorie,
armateur, plein-vide, états, POD, reefer, dangereux, dwell minimum, n° conteneur.

## Note sur le filtrage N4

L'endpoint `/apex/api/query?filtername=units_1` filtre selon un **filtre sauvegardé N4**
(les paramètres `operatorId/complexId/facilityId/yardId` sont des coordonnées de *scope*,
pas des filtres de données). Pour filtrer à la source, créer des filtres sauvegardés
dédiés dans N4 (`units_empty`, `units_export`…) et changer `filtername`. Le filtrage fin
se fait sinon côté serveur local, instantanément.
