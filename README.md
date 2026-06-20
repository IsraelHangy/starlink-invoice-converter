# Starlink to DEXY Converter

Application web React + TypeScript + Vite + Tailwind pour convertir un fichier Excel Starlink en template DEXY CD SFE.

## Fonction principale

La taxe `Congo DRC Telecommunication Excise Tax_Final` du fichier Starlink est transformée automatiquement en une seule ligne `TAX` par facture.

Pour chaque facture, l'application :

- conserve les lignes originales dans leur ordre ;
- additionne la taxe d'accise télécom ;
- ajoute une ligne `TAX` en fin de facture si le total est différent de 0 ;
- recalcule le numéro de ligne ;
- remplace `Date facture` par la date du jour de l'upload ;
- renseigne `Commentaire B` avec `Origine : [NUMERO_FACTURE] - [ANCIENNE_DATE]` ;
- exporte un fichier Excel officiel basé sur le template DEXY.

## Exports

- `Dexy_conforme_JJ-MM-AAAA.xlsx` : fichier officiel conforme DEXY, sans la colonne technique `Old Date`.

## Installation

```bash
pnpm install
pnpm run dev
```

## Build

```bash
pnpm run build
```

Le build statique est généré dans `dist/`.

## Déploiement Netlify

Le projet contient déjà `netlify.toml`.

Paramètres Netlify :

- Build command: `pnpm run build`
- Publish directory: `dist`

## Structure

```text
src/App.tsx
src/components/UploadCard.tsx
src/components/PreviewTable.tsx
src/components/SummaryCard.tsx
src/utils/converter.ts
src/utils/excel.ts
src/types.ts
```
