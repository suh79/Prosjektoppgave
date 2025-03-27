#Prosjektoppgave i PY1010 
#innlevert av Siri Undeland Haugen
# 27.03 2025

# Del a) Skriv et program som leser inn filen support_uke_24 og lagrer data
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# leser inn filen
fil='Prosjektoppgaven\support_uke_24.xlsx' #filnavnet 
data = pd.read_excel(fil) # leser inn filen

# Lagrer verdiene i Numpy arrays
u_dag = np.array(data['Ukedag'].values) # verdiene i kolonne 1
kl_slett = np.array(data['Klokkeslett'].values) # verdiene i kolonne 2
varighet = np.array(data['Varighet'].values) # verdiene i kolonne 3
score = np.array(data['Tilfredshet'].values) # verdiene i kolonne 4

# Skriv ut arrays for å bekrefte
print("Ukedag:", u_dag)
print("Klokkeslett:", kl_slett)
print("Varighet:", varighet)
print("Tilfredshet:", score)

#Del b) Skriv et program som finner antall henvendelser for hver av de 5 ukedagene. Resultatet 
#visualiseres ved bruk av et søylediagram (stolpediagram).

# tell antall for hver unike ukedag
unique, counts = np.unique(u_dag, return_counts=True)

# Lag en dictionary for å sortere ukedagene
ukedager_dict = {'Mandag': 0, 'Tirsdag': 1, 'Onsdag': 2, 'Torsdag': 3, 'Fredag': 4}
sorted_indices = np.argsort([ukedager_dict[dag] for dag in unique])

# Sorter ukedagene og counts i riktig rekkefølge
sorted_unique = unique[sorted_indices]
sorted_counts = counts[sorted_indices]

#vis ukedagene med antallet side ved side
print(np.asarray((sorted_unique, sorted_counts)).T)

# Lag et søylediagram
plt.figure(figsize=(10, 6))
plt.bar(sorted_unique, sorted_counts, color='skyblue')
plt.xlabel('Ukedag')
plt.ylabel('Antall henvendelser')
plt.title('Antall henvendelser per ukedag')
plt.show(block=False)  # Vis plottet uten å blokkere koden
plt.pause(10)  # Vent i 10 sekunder
plt.close()  # Lukk plottet

#Del c) Skriv et program som finner minste og lengste samtaletid som er loggført for uke 24. 
# svaret skrives til skjerm med informativ tekst.  

# Finn minste og lengste samtaletid som er loggført for uke 24
korteste_samtale = np.min(varighet)
lengste_samtale = np.max(varighet)
print("Minste samtaletid som er loggført for uke 24 varte i", korteste_samtale, "sekunder. Lengste samtaletid i samme uke varte i", lengste_samtale, "sekunder.")

# Del d) Beregn gjennomsnittlig samtaletid basert på alle henvendelser i uke 24
# Konverter varighet fra HH:MM:SS til sekunder
def tid_til_sekunder(tid):
    h, m, s = map(int, tid.split(':'))
    return h * 3600 + m * 60 + s

varighet_i_sekunder = np.array([tid_til_sekunder(tid) for tid in varighet])
#print("Varighet i sekunder:", varighet_i_sekunder)

# Konverter varighet fra sekunder til minutter
varighet_i_minutter = varighet_i_sekunder / 60
#print("Varighet i minutter:", varighet_i_minutter)

# Beregn gjennomsnittlig samtaletid i minutter
gjennomsnittlig_varighet_minutter = np.mean(varighet_i_minutter)
print(f"Gjennomsnittlig samtaletid for uke 24 er {gjennomsnittlig_varighet_minutter:.2f} minutter.")

#Del e) Supportvaktene i MORSE er delt inn i 2-timers bolker: kl 08-10, kl 10-12, kl 12-14 og kl 
#14-16. Skriv et program som finner det totale antall henvendelser supportavdelingen mottok 
#for hver av tidsrommene 08-10, 10-12, 12-14 og 14-16 for uke 24. Resultatet visualiseres ved 
#bruk av et sektordiagram (kakediagram). 

# Konverter klokkeslett fra HH:MM:SS til timer
def tid_til_timer(tid):
    h, m, s = map(int, tid.split(':'))
    return h + m / 60 + s / 3600

kl_slett_i_timer = np.array([tid_til_timer(tid) for tid in kl_slett])

# Del inn klokkeslettene i 2-timers bolker
bolker = np.digitize(kl_slett_i_timer, bins=[8, 10, 12, 14, 16])

# Tell antall henvendelser i hver bolk
bolk_counts = np.bincount(bolker)[1:]  # Ignorerer første element som er for verdier mindre enn første bin

# Definer etiketter for sektordiagrammet
etiketter = ['08-10', '10-12', '12-14', '14-16']

# Lag sektordiagram
plt.figure(figsize=(8, 8))
plt.pie(bolk_counts, labels=etiketter, autopct='%1.1f%%', startangle=140)
plt.title('Antall henvendelser per tidsrom for uke 24')
plt.show(block=False)  # Vis plottet uten å blokkere koden
plt.pause(10)  # Vent i 10 sekunder
plt.close()  # Lukk plottet

#Del f) Kundens tilfredshet loggføres som tall fra 1-10 hvor 1 indikerer svært misfornøyd og 
#10 indikerer svært fornøyd. Disse tilbakemeldingene skal så overføres til NPS-systemet (Net 
#Promoter Score).  
#NPS-systemet er konstruert på følgende måte: 
#Score 1-6 oppfattes som at kunden er negativ (vil trolig ikke anbefale MORSE til andre). 
#Score 7-8 oppfattes som et nøytralt svar. 
#Score 9-10 oppfattes som at kunden er positiv (vil trolig anbefale MORSE til andre).  
#Supportavdelingens NPS beregnes som et tall, prosentandelen positive kunder minus 
#prosentandelen negative kunder. Ved en formel kan dette gis slik: 
#NPS = % positive kunder - % negative kunder 
#Et eksempel på utregning av NPS er gitt i figuren under.  
#Kilde: https://www.blueprnt.com/2018/09/17/net-promoter-score/ 
#Lag et program som regner ut supportavdelings NPS og skriver svaret til skjerm. Merk: 
#Kunder som ikke har gitt tilbakemelding på tilfredshet, skal utelates fra utregningene.  

# Beregn supportavdelingens NPS
# Filtrer ut NaN-verdier fra score-arrayen
score_clean = score[~np.isnan(score)]

# Beregn antall positive, nøytrale og negative kunder
positive_kunder = np.sum((score_clean >= 9) & (score_clean <= 10))
noytrale_kunder = np.sum((score_clean >= 7) & (score_clean <= 8))
negative_kunder = np.sum((score_clean >= 1) & (score_clean <= 6))

# Beregn prosentandelene
total_kunder = len(score_clean)
prosent_positive = (positive_kunder / total_kunder) * 100
prosent_negative = (negative_kunder / total_kunder) * 100

# Beregn NPS
NPS = prosent_positive - prosent_negative
print(f"Supportavdelingens NPS for uke 24 er {NPS:.2f}")