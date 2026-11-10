# Modèle de calcul — Volumétrie réseau datacenter (multi-sites)

Classeur Excel à formules vivantes pour estimer les volumétries **entrantes / sortantes**
de datacenters structurés en PODs de 24 baies, en distinguant **internet (N-S)** et
**interne (E-O)**, avec consolidation multi-datacenters, trafic WAN inter-DC,
projection et dimensionnement.

## Fichiers

| Fichier | Rôle |
|---|---|
| `volumetrie_datacenter.xlsx` | Le modèle de calcul (ouvrir dans Excel / LibreOffice) |
| `generate_modele.py` | Script qui régénère le classeur (nécessite `openpyxl`) |

## Structure du classeur

1. **Mode d'emploi** — méthode, conversions, points de vigilance.
2. **Parametres** — simultanéité, ratio crête/moyenne, croissance, horizon, marge,
   sur-souscription, et noms des datacenters (DC1, DC2, DC3 — modifiables).
3. **Profils** — un profil par catégorie applicative : débit moyen/serveur (Mbps) et répartition
   du flux (sortant internet / entrant internet / interne LAN / réplication inter-DC).
   Les 4 parts de chaque profil doivent totaliser 100 % (colonne de contrôle).
4. **Inventaire** — une ligne par POD : datacenter (menu déroulant), salle, POD, nb baies,
   serveurs/baie, catégorie (menu déroulant). Nb de serveurs, débits et répartition par flux
   calculés automatiquement (formules posées jusqu'à la ligne 100).
5. **Resultats** — volumétrie globale par flux (moyenne, crête 95e centile approché, volumes
   mensuels en To), projection à l'horizon choisi, et capacité de dimensionnement recommandée
   (cœur internet / interne, uplink par POD).
6. **Synthese DC** — volumétrie par datacenter et global, trafic WAN inter-DC (décompte simple
   vs double comptage de la réplication), et scénario de bascule N+1 (datacenter défaillant et
   taux de bascule paramétrables).

## Conversions utilisées

- 1 Mbps soutenu × 1 mois (30 j) = **0,324 To**
- 1 Gb/s soutenu × 1 mois (30 j) = **324 To**

## Méthode (rappel)

```
Trafic POD (Mbps) = Σ (nb serveurs × débit moyen/serveur) × facteur de simultanéité
Crête POD         = trafic moyen × ratio crête/moyenne
Volume mensuel    = Mbps × 0,0108 × jours
Projection        = crête × (1 + croissance)^années × (1 + marge)
Bascule N+1       = crête propre du DC survivant + crête du DC défaillant × taux de bascule
```

## Multi-datacenters : conventions de comptage

- La **réplication inter-DC** est transportée aux deux extrémités : le total global la
  double-compte. L'onglet *Synthese DC* fournit le **volume unique** (décompte simple,
  /2) et le **WAN à provisionner** (décompte double).
- Le scénario de bascule dimensionne chaque site survivant : crête propre + crête du
  site défaillant × taux de bascule (PRA complet = 100 %).

## Calibration

Le modèle est **bottom-up** : calibrez-le sur vos mesures réelles (relevés 95e centile FAI,
compteurs SNMP du cœur de réseau, débits de réplication du stockage). Si le FAI indique
X Mbps sortants et le modèle Y, appliquez le coefficient Y/X au champ « Débit moyen / serveur »
des profils.

Si un site n'a pas de sortie internet locale (backhaul vers un site principal), réduisez la
part « sortant internet » de ses profils : son trafic est du WAN, pas de l'internet.

## Régénérer le classeur

```bash
pip install openpyxl
python3 generate_modele.py
```
