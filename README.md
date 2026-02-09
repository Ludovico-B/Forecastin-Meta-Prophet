# Forecast ML - Previsioni Ricavi Montagna 2026

Questo progetto utilizza il modello **Prophet** di Meta per analizzare i dati storici delle vendite e generare previsioni accurate dei ricavi per l'anno successivo

## 🚀 Funzionalità

- **Analisi Storica**: Caricamento e pulizia dei dati di vendita da file Excel.
- **Forecasting Avanzato**: Utilizzo di Prophet con stagionalità annuale, mensile e settimanale (moltiplicativa).
- **Visualizzazione Dati**: Generazione automatica di grafici dettagliati:
  - Forecast completo del periodo.
  - Componenti della stagionalità.
  - Focus specifico sull'anno 2026 con intervalli di confidenza.
  - Previsioni mensili aggregate per il 2026.
- **Esportazione Business**: Risultati salvati in un file Excel (`Previsioni_Prophet_2026.xlsx`) con riepiloghi dettagliati.

## 📁 Struttura del Progetto

- `prophet_forecasting.py`: Script principale per l'analisi e il forecasting.
- `explore_excel.py`: Utility per l'esplorazione rapida del dataset.
- `requirements.txt`: Elenco delle dipendenze Python.
- `Forecast immagini/`: Directory contenente i grafici generati.

## 🛠️ Installazione per Sviluppatori

Per far funzionare l'applicazione in locale, segui questi passaggi:

### 1. Prerequisiti
Assicurati di avere **Python 3.8+** installato sul tuo sistema.

### 2. Clonare il repository (o scaricare i file)
```bash
git clone https://github.com/tuo-username/forecast-ml.git
cd forecast-ml
```

### 3. Creare un ambiente virtuale (consigliato)
```bash
python -m venv venv
# Su Windows:
.\venv\Scripts\activate
# Su macOS/Linux:
source venv/bin/activate
```

### 4. Installare le dipendenze
```bash
pip install -r requirements.txt
```

## 📈 Come Eseguire l'App

1. Modificare e creare un file excel con Data di Vendita e Ricavo
2. Reimpostare la direcotiry del file excel e del foglio
3. Esegui lo script principale:
```bash
python prophet_forecasting.py
```

## 📊 Output Generati

Dopo l'esecuzione, troverai i seguenti file nella cartella di output:

- **Grafici (PNG)**:
  - `forecast_2026.png`: La curva di previsione per tutto il 2026.
  - `forecast_mensile_2026.png`: Istogramma dei ricavi previsti per mese.
  - `forecast_completo.png`: Visione d'insieme (storico + futuro).
- **Dati (Excel)**:
  - `Previsioni_Prophet_2026.xlsx`: Contiene le previsioni giornaliere, mensili e un foglio di riepilogo con le metriche chiave.

---
*Progetto sviluppato per l'analisi dei dati e il forecasting delle vendite.*
