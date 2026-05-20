# 🚀 DC Formatter - API et Pipeline de Transformation DOCX

> Un service containerisé pour transformer des documents DOCX à travers un pipeline complet d'extraction, conversion et rendu.

---

## 📋 Table des Matières

1. [Architecture](#architecture)
2. [Démarrage Rapide](#-démarrage-rapide)
3. [Utilisation de l'API](#-utilisation-de-lapi)
4. [Commandes Basiques](#-commandes-basiques)
5. [Configuration](#-configuration)
6. [Processus de Traitement](#-processus-de-traitement)
7. [Dépannage](#-dépannage)
8. [Déploiement](#-déploiement)

---

## Architecture

### Vue d'ensemble

```
┌─────────────────────────────────────────────────────────────┐
│                    🌐 Client (Navigateur/API)              │
│                  http://localhost:8000                      │
└──────────────────────────────┬────────────────────────────┘
                               │
                        HTTP POST /process
                               │
┌──────────────────────────────▼────────────────────────────┐
│                  🐳 Container Docker (dc-formatter)       │
│                                                            │
│  ┌────────────────────────────────────────────────────┐  │
│  │  FastAPI Application (app.py)                      │  │
│  │  ├─ Validation des fichiers                       │  │
│  │  ├─ Orchestration du pipeline                     │  │
│  │  ├─ Gestion des uploads/downloads                │  │
│  │  └─ Health checks                                │  │
│  └────────────────────────────────────────────────────┘  │
│                           │                                │
│                           ▼                                │
│  ┌────────────────────────────────────────────────────┐  │
│  │  DC Formatter Pipeline (tools/)                │  │
│  │  ├─ Stage 1: Extract XML from DOCX               │  │
│  │  ├─ Stage 2: Parse Template dimensions           │  │
│  │  ├─ Stage 3: XML → JSON (raw)                    │  │
│  │  ├─ Stage 4: Apply tags & styles (transformed)   │  │
│  │  └─ Stage 5: JSON → DOCX (final render)         │  │
│  └────────────────────────────────────────────────────┘  │
│                                                            │
│  Volumes (tous bindés):                                   │
│  • TEMPLATE ──────► ./app/TEMPLATE/ (template)          │
│  • DC_SOURCES ────► ./app/DC_SOURCES/ (input)           │
│  • OUTPUT1_XML-RAW ──► ./app/OUTPUT1_XML-RAW/           │
│  • OUTPUT2_JSON-RAW ──► ./app/OUTPUT2_JSON-RAW/         │
│  • OUTPUT3_JSON-TRANSFORMED ──► ./app/OUTPUT3_TRANSFORMED/ │
│  • OUTPUT4_DOCX-RESULT ──► ./app/OUTPUT4_DOCX-RESULT/   │
│  • OUTPUTS_FORMATTED ──► ./app/OUTPUTS_FORMATTED/       │
└──────────────────────────────┬────────────────────────────┘
                               │
                        HTTP Response (DOCX)
                               │
┌──────────────────────────────▼────────────────────────────┐
│              Client reçoit le fichier traité             │
└─────────────────────────────────────────────────────────┘
```

### Composants

| Composant | Type | Location | Rôle |
|-----------|------|----------|------|
| **FastAPI App** | Python | `/app/app.py` | API REST + orchestration |
| **Pipeline** | Python | `/app/tools/` | Logique de transformation |
| **Template** | DOCX | `/app/TEMPLATE/TEMPLATE.docx` | Template source (bindé) |
| **Entrypoint** | Bash | `/app/entrypoint.sh` | Fix permissions + démarrage |
| **Interface** | HTML | `/app/index.html` | Web UI simple |

---

## 🚀 Démarrage Rapide

### Option 1: Docker Compose (Recommandé)

```bash
# 1. Démarrer le service
docker-compose up -d

# 2. Vérifier que tout fonctionne
curl http://localhost:8000/health

# 3. Ouvrir l'interface web
# Navigateur: http://localhost:8000/
```

**Résultat attendu:**
```
{"status": "healthy", "service": "dc-formatter", "version": "1.0.0"}
```

### Option 2: Docker Seulement

```bash
# Build l'image
docker build -t dc-formatter:latest .

# Créer les répertoires de volume
mkdir -p app/{TEMPLATE,DC_SOURCES,OUTPUT1_XML-RAW,OUTPUT2_JSON-RAW,OUTPUT3_JSON-TRANSFORMED,OUTPUT4_DOCX-RESULT,OUTPUTS_FORMATTED}

# Lancer le container
docker run -d \
  --name dc-formatter \
  -p 8000:8000 \
  -v $(pwd)/app/TEMPLATE:/app/TEMPLATE \
  -v $(pwd)/app/DC_SOURCES:/app/DC_SOURCES \
  -v $(pwd)/app/OUTPUT1_XML-RAW:/app/OUTPUT1_XML-RAW \
  -v $(pwd)/app/OUTPUT2_JSON-RAW:/app/OUTPUT2_JSON-RAW \
  -v $(pwd)/app/OUTPUT3_JSON-TRANSFORMED:/app/OUTPUT3_JSON-TRANSFORMED \
  -v $(pwd)/app/OUTPUT4_DOCX-RESULT:/app/OUTPUT4_DOCX-RESULT \
  -v $(pwd)/app/OUTPUTS_FORMATTED:/app/OUTPUTS_FORMATTED \
  dc-formatter:latest

# Vérifier
curl http://localhost:8000/health
```

### Option 3: Développement (Python local)

```bash
# Créer virtualenv
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# ou: .venv\Scripts\activate  # Windows

# Installer dépendances
pip install -r requirements-api.txt

# Lancer l'app
python app.py
# Accédez à http://localhost:8000
```

---

## 📡 Utilisation de l'API

### Endpoints Disponibles

#### 1. Health Check

```bash
curl http://localhost:8000/health
```

Réponse:
```json
{
  "status": "healthy",
  "service": "dc-formatter",
  "version": "1.0.0"
}
```

#### 2. Traiter un Document (POST /process)

**Avec curl:**
```bash
curl -X POST http://localhost:8000/process \
  -F "file=@document.docx" \
  --output result.docx
```

**Avec Python:**
```python
import requests

url = "http://localhost:8000/process"
with open("document.docx", "rb") as f:
    response = requests.post(url, files={"file": f})
    
if response.ok:
    with open("result.docx", "wb") as out:
        out.write(response.content)
    print("✅ Document traité")
else:
    print(f"❌ Erreur: {response.json()}")
```

**Avec JavaScript/Node.js:**
```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

const form = new FormData();
form.append('file', fs.createReadStream('document.docx'));

axios.post('http://localhost:8000/process', form, {
  headers: form.getHeaders(),
  responseType: 'arraybuffer'
})
.then(response => {
  fs.writeFileSync('result.docx', response.data);
  console.log("✅ Document traité");
})
.catch(error => console.error("❌ Erreur:", error.message));
```

#### 3. Traiter Plusieurs Fichiers (POST /process-batch)

```bash
curl -X POST http://localhost:8000/process-batch \
  -F "files=@doc1.docx" \
  -F "files=@doc2.docx" \
  -F "files=@doc3.docx"
```

Réponse:
```json
{
  "results": [
    {"filename": "doc1.docx", "status": "success", "output_path": "..."},
    {"filename": "doc2.docx", "status": "success", "output_path": "..."},
    {"filename": "doc3.docx", "status": "error", "error": "..."}
  ]
}
```

#### 4. Documentation Interactive (Swagger UI)

```
http://localhost:8000/docs
```

Vous pouvez tester directement les endpoints depuis votre navigateur!

---

## ⌨️ Commandes Basiques

### Gestion du Service

```bash
# Démarrer
docker-compose up -d

# Arrêter
docker-compose down

# Redémarrer
docker-compose restart dc-formatter

# Voir les logs en temps réel
docker-compose logs -f dc-formatter

# Voir les 50 dernières lignes
docker-compose logs --tail=50
```

### Vérifier le Statut

```bash
# État du container
docker-compose ps

# Logs complets
docker-compose logs dc-formatter

# Statistiques (CPU, mémoire, I/O)
docker stats dc-formatter

# Health check manuel
curl http://localhost:8000/health
```

### Gestion des Fichiers

```bash
# Lister le template
ls -la app/TEMPLATE/

# Lister les fichiers sources
ls -la app/DC_SOURCES/

# Lister les fichiers pipeline intermédiaires
ls -la app/OUTPUT1_XML-RAW/
ls -la app/OUTPUT2_JSON-RAW/
ls -la app/OUTPUT3_JSON-TRANSFORMED/

# Lister les fichiers résultats
ls -la app/OUTPUT4_DOCX-RESULT/
ls -la app/OUTPUTS_FORMATTED/

# Copier un fichier depuis le container
docker cp dc-formatter:/app/OUTPUT4_DOCX-RESULT/file.docx ~/Desktop/

# Voir les logs détaillés
docker logs dc-formatter --tail=100
```

### Nettoyage

```bash
# Arrêter et supprimer le container
docker-compose down

# Supprimer aussi les volumes
docker-compose down -v

# Supprimer l'image
docker rmi dc-formatter:latest

# Nettoyer tout (containers, images, volumes)
docker system prune -a
```

---

## ⚙️ Configuration

### Variables d'Environnement

Ajouter dans `docker-compose.yml`:

```yaml
services:
  dc-formatter:
    environment:
      LOG_LEVEL: "INFO"              # DEBUG, INFO, WARNING, ERROR
      PYTHONUNBUFFERED: "1"
      PYTHONDONTWRITEBYTECODE: "1"
```

### Volumes

| Mount Point | Type | Host Path | Rôle |
|------------|------|-----------|------|
| `/app/TEMPLATE` | Bind | `./app/TEMPLATE/` | Template DOCX source |
| `/app/DC_SOURCES` | Bind | `./app/DC_SOURCES/` | Documents source (input API) |
| `/app/OUTPUT1_XML-RAW` | Bind | `./app/OUTPUT1_XML-RAW/` | XML bruts (stage 1 pipeline) |
| `/app/OUTPUT2_JSON-RAW` | Bind | `./app/OUTPUT2_JSON-RAW/` | JSON bruts (stage 2 pipeline) |
| `/app/OUTPUT3_JSON-TRANSFORMED` | Bind | `./app/OUTPUT3_JSON-TRANSFORMED/` | JSON transformés (stage 3 pipeline) |
| `/app/OUTPUT4_DOCX-RESULT` | Bind | `./app/OUTPUT4_DOCX-RESULT/` | DOCX résultats pipeline (stage 4) |
| `/app/OUTPUTS_FORMATTED` | Bind | `./app/OUTPUTS_FORMATTED/` | Fichiers finaux formatés (output app) |

### Ressources

```yaml
deploy:
  resources:
    limits:
      cpus: '2'              # Max 2 CPU
      memory: 2G             # Max 2 GB RAM
    reservations:
      cpus: '1'              # Réserver 1 CPU
      memory: 1G             # Réserver 1 GB RAM
```

### Port

Pour utiliser un port différent:

```yaml
services:
  dc-formatter:
    ports:
      - "8001:8000"  # Accès: http://localhost:8001
```

---

## 🔄 Processus de Traitement

### Pipeline Détaillé

Chaque document traverse 5 stages:

```
1️⃣  DOCUMENT DOCX (input)
        │
        ▼
    ┌─────────────────────┐
    │ Stage 1: Extract    │
    │ XML from DOCX       │
    │ (10+ fichiers XML)  │
    └─────────────────────┘
        │
        ▼
    🗂️  OUTPUT1_XML-RAW/ (temp)
        │
        ▼
    ┌─────────────────────┐
    │ Stage 2: Parse      │
    │ Template            │
    │ (dimensions, styles)│
    └─────────────────────┘
        │
        ▼
    ┌─────────────────────┐
    │ Stage 3: XML→JSON   │
    │ (raw conversion)    │
    └─────────────────────┘
        │
        ▼
    🗂️  OUTPUT2_JSON-RAW/ (temp)
        │
        ▼
    ┌─────────────────────┐
    │ Stage 4: Transform  │
    │ Apply tags & styles │
    │ (enrichissement)    │
    └─────────────────────┘
        │
        ▼
    🗂️  OUTPUT3_JSON-TRANSFORMED/ (temp)
        │
        ▼
    ┌─────────────────────┐
    │ Stage 5: Render     │
    │ JSON → DOCX        │
    │ (final document)    │
    └─────────────────────┘
        │
        ▼
    🗂️  OUTPUT4_DOCX-RESULT/ (output)
        │
        ▼
    📄 DOCUMENT DOCX (output)
        processed_filename.docx
```

### Fichiers de la Pipeline

Tous les dossiers OUTPUT sont bindés et persistés:

```python
# app.py - Pipeline stages
stage1_xml = export_all_xml(source_docx, OUTPUT1)           # → OUTPUT1_XML-RAW/
template_dims = parse_template(TEMPLATE_PATH)
stage2_json = xml_to_json(stage1_xml, OUTPUT2)              # → OUTPUT2_JSON-RAW/
stage3_transformed = apply_tags_and_styles(stage2_json, OUTPUT3, dims)  # → OUTPUT3_JSON-TRANSFORMED/
stage4_final = json_to_docx(stage3_transformed, TEMPLATE_PATH, OUTPUT4) # → OUTPUT4_DOCX-RESULT/
```

Après traitement:
- ✅ Tous les fichiers intermédiaires sont sauvegardés dans les dossiers OUTPUT correspondants
- ✅ Le fichier DOCX final est dans `OUTPUT4_DOCX-RESULT/`
- ✅ Les fichiers finaux formatés peuvent être placés dans `OUTPUTS_FORMATTED/`

### Temps de Traitement Typique

```
Stage 1 (Extract XML):      ~0.5-1.0s
Stage 2 (Parse Template):   ~0.01s
Stage 3 (XML→JSON):         ~0.2-0.3s
Stage 4 (Transform):        ~0.2-0.3s
Stage 5 (Render DOCX):      ~5.0-6.0s
                           ───────────
Total:                      ~6-8 secondes
```

---

## 🐛 Dépannage

### Erreur: Permission Denied (OUTPUT4_DOCX-RESULT)

**Cause:** Conflit de permissions entre l'utilisateur hôte et le user du container.

**Solution:**
```bash
# L'entrypoint.sh fixe automatiquement les permissions
# Si ça ne fonctionne pas, vérifier manuellement:

docker exec dc-formatter chmod 777 OUTPUT4_DOCX-RESULT
docker exec dc-formatter ls -la | grep OUTPUT
```

### Erreur: "Port 8000 already in use"

**Solution 1:** Utiliser un port différent
```yaml
# docker-compose.yml
ports:
  - "8001:8000"
```

**Solution 2:** Arrêter le service qui utilise le port
```bash
# Trouver le process
lsof -i :8000

# Ou simplement redémarrer le container
docker-compose down
docker-compose up -d
```

### Application ne démarre pas

**Vérifier les logs:**
```bash
docker-compose logs dc-formatter
```

**Causes courantes:**
- ❌ Fichier `entrypoint.sh` n'a pas les permissions d'exécution
  → Fixer: `chmod +x entrypoint.sh`

- ❌ Module `tools` non trouvé
  → Vérifier: `ls -la tools/`

- ❌ Dockerfile corrompu
  → Rebuild complet sans cache: `docker-compose down && docker build --no-cache -t dc-formatter:latest . && docker-compose up -d`

### API répond lentement

**Vérifier les ressources:**
```bash
# Voir la consommation
docker stats dc-formatter

# Augmenter les ressources dans docker-compose.yml
deploy:
  resources:
    limits:
      cpus: '4'      # Augmenter de 2 à 4
      memory: 4G     # Augmenter de 2G à 4G
```

### Fichiers output n'apparaissent pas

**Vérifier:**
```bash
# 1. Les répertoires existent?
ls -la app/OUTPUT*/ app/OUTPUTS_FORMATTED/

# 2. Les permissions sont bonnes?
chmod 777 app/OUTPUT*/ app/OUTPUTS_FORMATTED/

# 3. Vérifier dans le container
docker exec dc-formatter ls -la /app/OUTPUT4_DOCX-RESULT/
docker exec dc-formatter ls -la /app/OUTPUTS_FORMATTED/

# 4. Voir les logs
docker logs dc-formatter | tail -20
```

---

## 📦 Déploiement

### Docker Swarm

```bash
# Initialiser Swarm
docker swarm init

# Déployer le stack
docker stack deploy -c docker-compose.yml dc-formatter

# Voir le statut
docker stack services dc-formatter
```

### Kubernetes (avec Helm - optionnel)

```bash
# Créer un namespace
kubectl create namespace dc-formatter

# Déployer (si vous avez un chart Helm)
helm install dc-formatter ./helm-chart -n dc-formatter

# Vérifier
kubectl get pods -n dc-formatter
```

---

## 📚 Structure du Projet

```
dc_formatter/
├── app.py                         # FastAPI application
├── index.html                     # Web UI
├── entrypoint.sh                  # Startup script (permission fix)
├── Dockerfile                     # Container image
├── docker-compose.yml             # Orchestration
├── requirements-api.txt           # Python dependencies (for container)
├── README.md                      # Cette documentation
├── tools/                         # Pipeline modules
│   ├── extract_xml_raw.py
│   ├── parse_template.py
│   ├── parse_xml_raw_to_json_raw.py
│   ├── process_json_raw_to_json_transformed.py
│   └── render_json_transformed_to_docx.py
├── app/                           # Dossiers bindés du container
│   ├── TEMPLATE/                  # Template DOCX (bindé)
│   │   └── TEMPLATE.docx
│   ├── DC_SOURCES/                # Documents input (bindé)
│   ├── OUTPUT1_XML-RAW/           # XML bruts (bindé)
│   ├── OUTPUT2_JSON-RAW/          # JSON bruts (bindé)
│   ├── OUTPUT3_JSON-TRANSFORMED/  # JSON transformés (bindé)
│   ├── OUTPUT4_DOCX-RESULT/       # DOCX résultats (bindé)
│   └── OUTPUTS_FORMATTED/         # Résultats finaux (bindé)
├── .devcontainer/
│   └── devcontainer.json          # VS Code dev container
└── assets/                        # Fichiers statiques
    └── (images, ressources)
```

---

## 🔒 Sécurité

✅ **Container:**
- Non-root user (UID/GID 1000: `dcformatter`)
- Bash shell + home directory pour dev containers
- Nettoyage automatique des fichiers temporaires
- Health check pour le monitoring

✅ **API:**
- Validation stricte du type de fichier (.docx seulement)
- Gestion d'erreurs robuste avec logging
- Pas de secrets en dur (variables d'environnement)
- Timeout de traitement limité

✅ **Permissions:**
- Script `entrypoint.sh` fixe les permissions au démarrage
- Bind volumes avec ownership correct
- Volumes éphémères ignorés après traitement

---

## 📝 Notes de Version

- **Version:** 1.0.0
- **Base Image:** Python 3.11-slim
- **Framework:** FastAPI 0.104.1 + Uvicorn 0.24.0
- **Date:** Mai 2026

---

## 🆘 Support

Pour les problèmes:

1. **Vérifier les logs:**
   ```bash
   docker-compose logs -f dc-formatter
   ```

2. **Vérifier la santé:**
   ```bash
   curl http://localhost:8000/health
   ```

3. **Tester l'API:**
   ```bash
   curl -X POST http://localhost:8000/process \
     -F "file=@test.docx" \
     --output test_result.docx
   ```

4. **Vérifier les ressources:**
   ```bash
   docker stats dc-formatter
   ```

---

## 🖥️ Utilisation en Ligne de Commande (Pipeline CLI)

### Installation Locale

```bash
pip install -r requirements.txt
```

### Pipeline Complet (CLI)

```bash
python3 -m tools.pipeline full -s DC_JNZ_2026.docx
```

### Résultats
Les fichiers générés sont organisés dans 4 dossiers:
- `OUTPUT1_XML-RAW/` → XML global brut
- `OUTPUT2_JSON-RAW/` → Structure JSON complète
- `OUTPUT3_JSON-TRANSFORMED/` → JSON avec tags et styles
- `OUTPUT4_DOCX-RESULT/` → Document Word final

### 🛠️ Modules Python (tools/)

| Module | Fonction |
|--------|----------|
| `extract_xml_raw.py` | Extraction XML depuis le DOCX (OUTPUT1) |
| `parse_template.py` | Extraction dimensions et paramètres du template |
| `parse_xml_raw_to_json_raw.py` | Conversion XML → JSON brut (OUTPUT2) |
| `process_json_raw_to_json_transformed.py` | Application tags/styles (OUTPUT3) |
| `render_json_transformed_to_docx.py` | Génération DOCX final (OUTPUT4) |
| `pipeline.py` | Interface CLI (orchestration) |
| `zip_docx.py` | Archivage et compression |

### 📊 Structure des Données

```
DOCX (zippé)
    ↓ extraction
XML brut (word/document.xml + relations)
    ↓ parsing + structuration
JSON RAW (paragraphes, tables, runs bruts)
    ↓ tagging + détection hiérarchie
JSON TRANSFORMED (avec structures sections + tags)
    ↓ rendu template
DOCX final (formaté + stylisé)
```

### 📚 Documentation Détaillée

Pour plus d'informations sur le pipeline CLI et l'architecture avancée:

Voir [Z_DOCUMENTATION/README.md](Z_DOCUMENTATION/README.md) pour:
- Architecture détaillée du pipeline
- Description de chaque module Python
- Analyse des formats (XML, JSON, DOCX)
- Guide de détection des hiérarchies
- Commandes avancées

## 🛠️ Modules Python

| Module | Fonction |
|--------|----------|
| `extract_xml_raw.py` | Extraction XML depuis le DOCX |
| `parse_template.py` | Extraction dimensions et paramètres |
| `parse_xml_raw_to_json_raw.py` | Conversion XML → JSON brut |
| `process_json_raw_to_json_transformed.py` | Application tags/styles |
| `render_json_transformed_to_docx.py` | Génération DOCX final |
| `pipeline.py` | Interface CLI (orchestration) |
| `zip_docx.py` | Archivage et compression |

## 📊 Structure des Données

```
DOCX (zippé)
    ↓ extraction
XML brut (word/document.xml + relations)
    ↓ parsing + structuration
JSON RAW (paragraphes, tables, runs bruts)
    ↓ tagging + détection hiérarchie
JSON TRANSFORMED (avec structures sections + tags)
    ↓ rendu template
DOCX final (formaté + stylisé)
```

## 📝 Exemples Couramment Utilisés

```bash
# Pipeline complet
python3 -m tools.pipeline full -s document.docx

# Avec output custom
python3 -m tools.pipeline full -s document.docx -o results/

# Phase 1 seulement (extraction)
python3 -m tools.pipeline extract -s document.docx

# Phase 2 seulement (transformation)
python3 -m tools.pipeline transform-render -s document.docx

# Extraire dimensions du template
python3 -m tools.pipeline extract-dims
```

## 🔧 Configuration

- **Template par défaut**: `TEMPLATE/TEMPLATE.docx`
- **Sources DOCX**: Placez les fichiers dans `DC_SOURCES/` ou utilisez `-s chemin/complet`
- **Dossier sortie**: Paramètre `-o` (par défaut: OUTPUT*/)

## ⚡ Toutes les Commandes

Voir [Z_DOCUMENTATION/COMMANDES.sh](Z_DOCUMENTATION/COMMANDES.sh) pour la liste exhaustive avec exemples.
