# Sensor Dashboard — Dwarsprofiel Uden

Interactief dashboard voor temperatuur en luchtvochtigheid van sensoren op verschillende hoogtes.

## Functie

- Laadt live data uit het Excel-bestand op `water-tech.cboost.nl/excel`
- Toont daggrafieken per sensor (VDB45, VDB39, VDB37, VDB36, VDB14)
- Navigeer per dag met de ‹ › knoppen of pijltoetsen
- Schakel tussen temperatuur en luchtvochtigheid
- Data wordt 20 minuten gecached

## Deployen op Render

1. Fork of push dit project naar GitHub
2. Ga naar [render.com](https://render.com) → **New Web Service**
3. Koppel de GitHub repository
4. Instellingen:
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Environment Variable:** `EXCEL_URL = https://water-tech.cboost.nl/excel`
5. Deploy

## Lokaal draaien

```bash
npm install
EXCEL_URL=https://water-tech.cboost.nl/excel npm start
```

Open dan `http://localhost:3000`

## Sensorhoogtes

| Sensor | Hoogte |
|--------|--------|
| VDB45  | 29 cm  |
| VDB39  | 100 cm |
| VDB37  | 130 cm |
| VDB36  | 207 cm |
| VDB14  | 283 cm |

