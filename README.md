# Modèle de calcul — Volumétrie réseau datacenter

Classeur Excel à formules vivantes pour estimer les volumétries **entrantes / sortantes**
d'un datacenter structuré en PODs de 24 baies, en distinguant **internet (N-S)** et
**interne (E-O)**, avec projection et dimensionnement.

## Fichiers

| Fichier | Rôle |
|---|---|
| `volumetrie_datacenter.xlsx` | Le modèle de calcul (ouvrir dans Excel / LibreOffice) |
| `generate_modele.py` | Script qui régénère le classeur (nécessite `openpyxl`) |

## Structure du classeur

1. **Mode d'emploi** — méthode, conversions, points de vigilance.
2. **Parametres** — simultanéité, ratio crête/moyenne, croissance, horizon, marge, sur-souscription.
3. **Profils** — un profil par catégorie applicative : débit moyen/serveur (Mbps) et répartition
   du flux (sortant internet / entrant internet / interne LAN / réplication inter-DC).
   Les 4 parts de chaque profil doivent totaliser 100 % (colonne de contrôle).
4. **Inventaire** — une ligne par POD : salle, POD, nb baies, serveurs/baie, catégorie (menu
   déroulant). Nb de serveurs, débits et répartition par flux calculés automatiquement
   (formules posées jusqu'à la ligne 100).
5. **Resultats** — volumétrie actuelle par flux (moyenne, crête 95e centile approché, volumes
   mensuels en To), projection à l'horizon choisi, et capacité de dimensionnement recommandée
   (cœur internet / interne, uplink par POD).

## Conversions utilisées

- 1 Mbps soutenu × 1 mois (30 j) = **0,324 To**
- 1 Gb/s soutenu × 1 mois (30 j) = **324 To**

## Méthode (rappel)

```
Trafic POD (Mbps) = Σ (nb serveurs × débit moyen/serveur) × facteur de simultanéité
Crête POD         = trafic moyen × ratio crête/moyenne
Volume mensuel    = Mbps × 0,0108 × jours
Projection        = crête × (1 + croissance)^années × (1 + marge)
```

## Calibration

Le modèle est **bottom-up** : calibrez-le sur vos mesures réelles (relevés 95e centile FAI,
compteurs SNMP du cœur de réseau, débits de réplication du stockage). Si le FAI indique
X Mbps sortants et le modèle Y, appliquez le coefficient Y/X au champ « Débit moyen / serveur »
des profils.

## Régénérer le classeur

```bash
pip install openpyxl
python3 generate_modele.py
```
