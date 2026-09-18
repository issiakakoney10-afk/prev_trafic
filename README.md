# Outil local — Prévisions de trafics

Application locale (Streamlit) qui régénère votre fichier Excel de
prévisions de trafics, **au même format qu'aujourd'hui**, à partir de
données mises à jour. Les formules d'origine ne sont jamais réécrites :
l'outil ne fait que remplir les cellules de saisie, puis recalcule tout le
classeur, exactement comme Excel le ferait.

Tout se passe sur votre ordinateur : aucune donnée n'est envoyée sur
Internet.

## 1. Prérequis

- **Python 3.10 ou plus récent**.
- **LibreOffice** installé (fournit la commande `soffice`, utilisée en
  arrière-plan pour recalculer le classeur final — vous n'avez pas besoin
  d'ouvrir LibreOffice vous-même, ni même de savoir qu'il est là) :
  - **Windows** : téléchargez et installez LibreOffice depuis
    [libreoffice.org](https://www.libreoffice.org/download/download/). La
    commande `soffice` est installée automatiquement.
  - **macOS** : `brew install --cask libreoffice`, ou installez depuis
    [libreoffice.org](https://www.libreoffice.org/download/download/).
  - **Linux (Debian/Ubuntu)** : `sudo apt install libreoffice`.

Vérifiez l'installation avec :

```bash
soffice --version
```

## 2. Installation

Depuis le dossier du projet :

```bash
pip install -r requirements.txt
```

(Sur certains systèmes, utilisez `pip3` à la place de `pip`.)

## 3. Lancer l'outil

Toujours depuis le dossier du projet :

```bash
streamlit run app.py
```

Une page s'ouvre dans votre navigateur (sinon, l'adresse à ouvrir
manuellement, en général `http://localhost:8501`, s'affiche dans le
terminal). Pour arrêter l'outil, retournez dans le terminal et faites
`Ctrl+C`.

## 4. Utilisation

L'application est organisée en 5 onglets :

### 📅 Saisie mensuelle

En haut, téléchargez le modèle vierge **Détail mensuel 2026** (navires,
tonnages par produit, conteneurs...). Ouvrez-le dans Excel — **seules les
cellules surlignées en jaune sont modifiables** (le reste est protégé), ce
sont exactement les mêmes cellules que vous rempliriez dans le fichier
d'origine. Redéposez-le une fois rempli (optionnel : ne déposez que ce que
vous avez à mettre à jour ce jour-là).

En dessous, le **chiffre annoncé du mois** (optionnel) : pour un mois dont
le détail complet (navire par navire) n'est pas encore saisi, deux chiffres
"flash" suffisent — le nombre d'escales de navires et le tonnage global du
mois. Le fichier estime alors automatiquement les autres statistiques
annoncées (national, transit, transbordement, conteneurs...) à partir de
ces deux chiffres, exactement comme il le fait déjà aujourd'hui dans votre
classeur.

### 📊 Coefficients saisonniers

Le modèle vierge **Paramètres annuels** (objectifs annuels, historique 5
ans, engagement contractuel TC2, statistiques de clôture) se télécharge et
se redépose ici, séparément de la génération des prévisions. C'est cette
saisonnalité qui alimente la méthode « coefficients saisonniers » de
l'onglet suivant — d'où la séparation : on renseigne l'hypothèse de fond
ici, on choisit comment l'utiliser dans Prévisions. Un aperçu interactif
(coefficient par mois, historique 5 ans) est affiché en bas de l'onglet.

### 🔮 Prévisions

Deux réglages à ne pas confondre :

- **Dernier mois en détail complet** — pilote la feuille **CALC. PREV**.
  Jusqu'à ce mois inclus, CALC. PREV utilise le détail complet saisi
  (feuille "Trafic mens"). Au-delà, elle utilise le chiffre annoncé du mois
  (onglet Saisie mensuelle), s'il a été renseigné.
- **Dernier mois connu** — pilote la feuille **ANNUEL**, séparément.
  Jusqu'à ce mois inclus, ANNUEL reprend directement le réalisé de
  CALC. PREV (détail complet ou chiffre annoncé, peu importe). Au-delà,
  ANNUEL bascule sur la **méthode de projection** choisie juste en dessous,
  pour estimer les mois restants de l'année.

En général, "dernier mois connu" est égal ou postérieur à "dernier mois en
détail complet" (on peut "connaître" un mois via son seul chiffre annoncé,
avant d'en avoir le détail complet).

**Important** : chaque génération repart du fichier d'origine. Ces réglages
ne sont pas mémorisés d'une génération à l'autre — si la frontière doit
avancer d'un mois par rapport à la dernière fois, il faut la ressaisir.

Cliquez sur **Générer le fichier Excel**. Le recalcul complet peut prendre
jusqu'à une minute. Une fois terminé, téléchargez le fichier avec le bouton
prévu à cet effet (il reste aussi disponible dans le dossier `output/` du
projet).

### 📈 Visualisation

Explore, en tableaux et graphiques interactifs, le dernier fichier généré
dans la session en cours (ou le fichier de référence tant que rien n'a
encore été généré) : vue d'ensemble ANNUEL (réalisé en marqueur plein,
projeté en marqueur creux), CALC. PREV, détail mensuel par année,
répartition pluriannuelle, engagement TC2, statistiques de clôture.

### ❓ Aide

Résumé du fonctionnement, du lien entre les onglets et des erreurs
courantes — directement dans l'application.

## 5. Structure du projet

```
trafic-tool/
├── app.py                     interface (à lancer avec streamlit run)
├── requirements.txt
├── README.md
├── schema/
│   └── master_reference.xlsx  fichier maître d'origine — ne jamais modifier
├── templates/                 modèles vierges téléchargeables depuis l'appli
├── output/                    fichiers générés (peut être vidé à tout moment)
└── engine/                    moteur de génération (extraction, modèles, injection)
```

## 6. Limites connues

- **Feuille `ANNUEL_DCAQ_AVEC_TC2_Option2`** : ancienne version (2021-2023),
  contient des erreurs `#REF!` héritées. Elle est recopiée telle quelle dans
  chaque fichier généré, sans aucun recalcul — comme dans le fichier
  d'origine.
- **Roulement annuel non automatisé** : passer à une nouvelle année
  (nouvel onglet "Trafic mens", décalage de l'historique 5 ans) n'est pas
  encore automatisé. L'historique 5 ans reste modifiable à la main via le
  modèle "Paramètres annuels" (onglet COEFFICIENTS_SAISONNIERS_2026). Le
  jour où un véritable nouvel onglet annuel est nécessaire, l'outil devra
  être adapté au nouveau fichier de référence.
- **Le fichier de référence est fixe** : l'outil part toujours de
  `schema/master_reference.xlsx`. Si la structure du classeur d'origine
  change un jour (lignes ajoutées/supprimées, feuilles renommées...), ce
  fichier de référence doit être remplacé et l'outil revérifié en
  conséquence.

## 7. En cas de problème

- **"LibreOffice (commande 'soffice') n'a pas été trouvé"** : LibreOffice
  n'est pas installé, ou pas dans le PATH — revoir la section 1.
  Redémarrez le terminal après l'installation.
- **Le port 8501 est déjà utilisé** : lancez `streamlit run app.py
  --server.port 8502` (ou un autre numéro), puis ouvrez l'adresse indiquée
  dans le terminal.
- **Erreur pendant la génération** : le message affiché indique la cause
  (le plus souvent, un modèle déposé ne correspond pas à ce qui est
  attendu — redéposez un modèle téléchargé depuis l'outil lui-même).

## 8. Déploiement en ligne (sans rien installer localement)

Si vous ne pouvez pas installer Python/LibreOffice sur votre poste, ce
projet peut aussi tourner comme un site accessible par un lien, hébergé
gratuitement sur **Streamlit Community Cloud**. Guide détaillé fourni à
part ; résumé des étapes :

1. Créez un compte GitHub gratuit et un **dépôt privé** (important : les
   données de prévision sont internes — ne pas utiliser un dépôt public).
2. Déposez-y le contenu de ce dossier via l'upload web de GitHub
   (glisser-déposer dans le navigateur, aucune ligne de commande requise).
   Ce dépôt contient déjà `packages.txt` (installe LibreOffice côté
   serveur).
3. Sur [share.streamlit.io](https://share.streamlit.io), connectez-vous
   avec votre compte GitHub, puis cliquez sur **Create app**. À la question
   « Do you already have an app? », choisissez **Yup, I have an app**. Le
   plus simple : cliquez sur **Paste GitHub URL** et collez le lien vers
   `app.py` dans votre dépôt — sinon, renseignez le dépôt, la branche
   (`main`) et le chemin `app.py` séparément. Ouvrez **Advanced settings**
   et choisissez la version Python **3.11**, puis **Deploy**.
4. Une fois en ligne, cliquez sur **Share** (en haut à droite), désactivez
   l'option qui rend l'app publique, et ajoutez les adresses e-mail des
   personnes autorisées (ou : réglages de l'app → **Sharing** → **Only
   specific people can view this app**). Le compte gratuit n'autorise
   qu'une seule application privée à la fois.

Le moteur est déjà prêt pour un usage à plusieurs (chaque génération
utilise un profil LibreOffice et des fichiers temporaires isolés, pour
rester correcte si deux personnes cliquent sur "Générer" en même temps).
