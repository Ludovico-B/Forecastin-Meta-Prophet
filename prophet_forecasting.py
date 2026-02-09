"""
Forecasting delle Vendite con Prophet
Analisi del file: Analisi Montagna Forecasting 2026 (dal 2024).xlsx
Foglio: VEND - Colonne: Data Vendita, Ricavo
"""

import pandas as pd
import numpy as np
from prophet import Prophet
import matplotlib.pyplot as plt
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# Configurazione matplotlib
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['figure.figsize'] = (14, 8)
plt.rcParams['font.size'] = 12

# 1. Caricamento dei dati
print("=" * 60)
print("FORECASTING VENDITE MONTAGNA CON PROPHET")
print("=" * 60)

file_path = r"c:\Users\Vstor\Desktop\Ludovico\Analisi Dati\Previsioni 2026 Montagna\Analisi Montagna Forecasting 2026 (dal 2024).xlsx"
output_dir = r"c:\Users\Vstor\Desktop\Ludovico\Analisi Dati\Previsioni 2026 Montagna"

print("\n📊 Caricamento dati dal foglio VEND...")
df = pd.read_excel(file_path, sheet_name='VEND')

print(f" Numero totale righe: {len(df)}")

# 2. Preparazione dei dati per Prophet
print("\n🔧 Preparazione dati per Prophet...")

# Crea DataFrame per Prophet con le colonne corrette
df_prophet = df[['Data Vendita', 'Ricavo']].copy()
df_prophet.columns = ['ds', 'y']

# Pulizia dati
df_prophet['ds'] = pd.to_datetime(df_prophet['ds'], errors='coerce')
df_prophet['y'] = pd.to_numeric(df_prophet['y'], errors='coerce')
df_prophet = df_prophet.dropna()

# Aggrega per giorno
df_daily = df_prophet.groupby('ds').agg({'y': 'sum'}).reset_index()
df_daily = df_daily.sort_values('ds')

print(f"\n📅 Range dati: {df_daily['ds'].min().strftime('%Y-%m-%d')} → {df_daily['ds'].max().strftime('%Y-%m-%d')}")
print(f"📊 Giorni con dati: {len(df_daily)}")
print(f"💰 Ricavo totale: €{df_daily['y'].sum():,.2f}")
print(f"💰 Ricavo medio giornaliero: €{df_daily['y'].mean():,.2f}")

# 3. Statistiche descrittive
print("\n" + "=" * 60)
print("STATISTICHE DESCRITTIVE")
print("=" * 60)
print(f"   Minimo:   €{df_daily['y'].min():,.2f}")
print(f"   Massimo:  €{df_daily['y'].max():,.2f}")
print(f"   Media:    €{df_daily['y'].mean():,.2f}")
print(f"   Mediana:  €{df_daily['y'].median():,.2f}")
print(f"   Std Dev:  €{df_daily['y'].std():,.2f}")

# 4. Creazione e training del modello Prophet
print("\n" + "=" * 60)
print("TRAINING MODELLO PROPHET")
print("=" * 60)

model = Prophet(
    yearly_seasonality=True,
    weekly_seasonality=True,
    daily_seasonality=False,
    seasonality_mode='multiplicative',
    changepoint_prior_scale=0.05,
    seasonality_prior_scale=10,
)

model.add_seasonality(name='monthly', period=30.5, fourier_order=5)

print("🤖 Training del modello in corso...")
model.fit(df_daily)
print("✅ Modello addestrato con successo!")

# 5. Creazione previsioni per il 2026
print("\n" + "=" * 60)
print("GENERAZIONE PREVISIONI 2026")
print("=" * 60)

last_date = df_daily['ds'].max()
end_2026 = pd.Timestamp('2026-12-31')
days_to_predict = (end_2026 - last_date).days

if days_to_predict <= 0:
    future = model.make_future_dataframe(periods=365)
else:
    future = model.make_future_dataframe(periods=days_to_predict)

print(f"📆 Previsioni generate fino al: {future['ds'].max().strftime('%Y-%m-%d')}")

forecast = model.predict(future)

# 6. Filtra solo le previsioni per il 2026
forecast_2026 = forecast[forecast['ds'].dt.year == 2026].copy()
print(f"📊 Giorni previsti per il 2026: {len(forecast_2026)}")

print(f"\n💰 PREVISIONI RICAVI 2026:")
print(f"   Ricavo totale previsto:  €{forecast_2026['yhat'].sum():,.2f}")
print(f"   Ricavo medio giornaliero: €{forecast_2026['yhat'].mean():,.2f}")
print(f"   Intervallo confidenza (lower): €{forecast_2026['yhat_lower'].sum():,.2f}")
print(f"   Intervallo confidenza (upper): €{forecast_2026['yhat_upper'].sum():,.2f}")

# Previsioni mensili 2026
print("\n📅 PREVISIONI MENSILI 2026:")
forecast_2026['month'] = forecast_2026['ds'].dt.month
monthly_forecast = forecast_2026.groupby('month').agg({
    'yhat': 'sum',
    'yhat_lower': 'sum',
    'yhat_upper': 'sum'
}).reset_index()

mesi_italiani = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
                 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre']

print(f"\n{'Mese':<12} {'Previsto':>15} {'Min':>15} {'Max':>15}")
print("-" * 60)
for _, row in monthly_forecast.iterrows():
    mese = mesi_italiani[int(row['month']) - 1]
    print(f"{mese:<12} €{row['yhat']:>13,.2f} €{row['yhat_lower']:>13,.2f} €{row['yhat_upper']:>13,.2f}")

# 7. Creazione grafici
print("\n" + "=" * 60)
print("GENERAZIONE GRAFICI")
print("=" * 60)

# Grafico 1: Forecast completo
fig1 = model.plot(forecast)
plt.title('Forecasting Ricavi Montagna - Prophet', fontsize=16, fontweight='bold')
plt.xlabel('Data', fontsize=12)
plt.ylabel('Ricavo (€)', fontsize=12)
plt.tight_layout()
fig1.savefig(f"{output_dir}/forecast_completo.png", dpi=150, bbox_inches='tight')
print("✅ Salvato: forecast_completo.png")
plt.close(fig1)

# Grafico 2: Componenti del modello
fig2 = model.plot_components(forecast)
plt.tight_layout()
fig2.savefig(f"{output_dir}/forecast_componenti.png", dpi=150, bbox_inches='tight')
print("✅ Salvato: forecast_componenti.png")
plt.close(fig2)

# Grafico 3: Focus sul 2026
fig3, ax = plt.subplots(figsize=(16, 8))

historical = df_daily[df_daily['ds'].dt.year < 2026]
if len(historical) > 0:
    ax.scatter(historical['ds'], historical['y'], alpha=0.3, s=10, label='Dati Storici', color='blue')

ax.plot(forecast_2026['ds'], forecast_2026['yhat'], color='red', linewidth=2, label='Previsione 2026')
ax.fill_between(forecast_2026['ds'], forecast_2026['yhat_lower'], forecast_2026['yhat_upper'], 
                alpha=0.3, color='red', label='Intervallo Confidenza')

ax.set_title('Previsione Ricavi 2026 - Montagna', fontsize=16, fontweight='bold')
ax.set_xlabel('Data', fontsize=12)
ax.set_ylabel('Ricavo (€)', fontsize=12)
ax.legend(loc='upper left')
ax.grid(True, alpha=0.3)
plt.tight_layout()
fig3.savefig(f"{output_dir}/forecast_2026.png", dpi=150, bbox_inches='tight')
print("✅ Salvato: forecast_2026.png")
plt.close(fig3)

# Grafico 4: Previsioni mensili 2026
fig4, ax = plt.subplots(figsize=(14, 8))
x = range(len(monthly_forecast))
bars = ax.bar(x, monthly_forecast['yhat'], color='steelblue', alpha=0.8, edgecolor='navy')

yerr_lower = monthly_forecast['yhat'] - monthly_forecast['yhat_lower']
yerr_upper = monthly_forecast['yhat_upper'] - monthly_forecast['yhat']
ax.errorbar(x, monthly_forecast['yhat'], yerr=[yerr_lower, yerr_upper], 
            fmt='none', color='darkred', capsize=5, capthick=2)

ax.set_xticks(x)
ax.set_xticklabels(mesi_italiani, rotation=45, ha='right')
ax.set_title('Previsione Ricavi Mensili 2026 - Montagna', fontsize=16, fontweight='bold')
ax.set_xlabel('Mese', fontsize=12)
ax.set_ylabel('Ricavo Previsto (€)', fontsize=12)

for i, (bar, val) in enumerate(zip(bars, monthly_forecast['yhat'])):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + yerr_upper.iloc[i] + 1000, 
            f'€{val/1000:.1f}k', ha='center', va='bottom', fontsize=9, fontweight='bold')

plt.tight_layout()
fig4.savefig(f"{output_dir}/forecast_mensile_2026.png", dpi=150, bbox_inches='tight')
print("✅ Salvato: forecast_mensile_2026.png")
plt.close(fig4)

# 8. Esporta risultati in Excel
print("\n" + "=" * 60)
print("ESPORTAZIONE RISULTATI")
print("=" * 60)

export_daily = forecast_2026[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].copy()
export_daily.columns = ['Data', 'Ricavo_Previsto', 'Ricavo_Min', 'Ricavo_Max']
export_daily['Data'] = export_daily['Data'].dt.strftime('%Y-%m-%d')

export_monthly = monthly_forecast.copy()
export_monthly['Mese'] = export_monthly['month'].apply(lambda x: mesi_italiani[int(x) - 1])
export_monthly = export_monthly[['Mese', 'yhat', 'yhat_lower', 'yhat_upper']]
export_monthly.columns = ['Mese', 'Ricavo_Previsto', 'Ricavo_Min', 'Ricavo_Max']

with pd.ExcelWriter(f"{output_dir}/Previsioni_Prophet_2026.xlsx", engine='openpyxl') as writer:
    export_daily.to_excel(writer, sheet_name='Previsioni_Giornaliere', index=False)
    export_monthly.to_excel(writer, sheet_name='Previsioni_Mensili', index=False)
    
    summary = pd.DataFrame({
        'Metrica': ['Ricavo Totale Previsto 2026', 'Ricavo Medio Giornaliero', 
                    'Intervallo Minimo', 'Intervallo Massimo', 
                    'Data Analisi', 'Modello Utilizzato'],
        'Valore': [f"€{forecast_2026['yhat'].sum():,.2f}", 
                   f"€{forecast_2026['yhat'].mean():,.2f}",
                   f"€{forecast_2026['yhat_lower'].sum():,.2f}", 
                   f"€{forecast_2026['yhat_upper'].sum():,.2f}",
                   datetime.now().strftime('%Y-%m-%d %H:%M'),
                   'Facebook Prophet (Meta)']
    })
    summary.to_excel(writer, sheet_name='Riepilogo', index=False)

print("✅ Salvato: Previsioni_Prophet_2026.xlsx")

print("\n" + "=" * 60)
print("✅ ANALISI COMPLETATA CON SUCCESSO!")
print("=" * 60)
print(f"\n📁 File generati in: {output_dir}")
print("   • forecast_completo.png")
print("   • forecast_componenti.png")
print("   • forecast_2026.png")
print("   • forecast_mensile_2026.png")
print("   • Previsioni_Prophet_2026.xlsx")
