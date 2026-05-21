# DeFi Monitoring - Données on-chain en temps réel

Script de monitoring des marchés de prêt USDC/USDT sur Morpho Blue et Aave V3, lisant les données directement depuis les smart contracts Ethereum via Alchemy.

## Ce que fait le script

- Lit les APY, TVL, taux d'utilisation et liquidité disponible directement on-chain (sans agrégateur tiers)
- Surveille les marchés actifs : Morpho Blue WBTC/USDC, Morpho Blue WETH/USDC, Morpho Blue wstETH/USDT, Aave V3 USDC
- Se rafraîchit automatiquement toutes les 5 minutes
- Exporte les données dans un fichier Excel avec deux onglets : dernière lecture et historique complet
- Génère des alertes si l'utilisation dépasse 90% (risque de blocage des retraits) ou si l'APY dépasse le seuil configuré

## Installation

Python 3.8+ requis.

```
pip install web3 python-dotenv openpyxl
```

Créer un compte gratuit sur [alchemy.com](https://alchemy.com), créer une app Ethereum Mainnet, et copier l'URL de l'API.

Créer un fichier `.env` dans le même dossier que le script :

```
ALCHEMY_URL=https://eth-mainnet.g.alchemy.com/v2/TA_CLE_ICI
```

## Lancement

Lecture unique :
```
python3 defi_monitoring.py once
```

Mode continu, toutes les 5 minutes :
```
python3 defi_monitoring.py
```

Arrêter avec Ctrl+C.

## Fichier Excel généré

Le fichier `monitoring_defi.xlsx` est créé automatiquement au premier lancement :

- **Dernière lecture** : snapshot du cycle en cours, mis à jour à chaque cycle
- **Historique** : toutes les observations accumulées, jamais écrasées

## Sources de données

- Marchés Morpho Blue : lecture via `market()` et `borrowRateView()` directement sur le contrat Morpho Blue (`0xBBBBBbb...`), paramètres récupérés automatiquement via `idToMarketParams()`
- Aave V3 : lecture via `getReserveData()` sur le contrat PoolDataProvider
- APY calculé depuis les borrow rates on-chain par composition continue annualisée

## Stack

Python, web3.py, Alchemy, openpyxl
