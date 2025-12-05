# MultiZip - Estrattore ZIP Avanzato per Windows

MultiZip è uno **script PowerShell con interfaccia grafica (GUI)** per estrarre facilmente file ZIP su Windows.  
Permette di selezionare una cartella sorgente e una destinazione, gestire più ZIP contemporaneamente, e offre opzioni avanzate come mantenere la struttura delle sottocartelle, sovrascrivere file esistenti o eliminare gli ZIP originali dopo l’estrazione.

---

## 📌 Caratteristiche principali

- Selezione grafica di **cartella sorgente** e **cartella destinazione**  
- Estrazione di **più file ZIP** presenti nella cartella sorgente  
- Opzioni configurabili:
  - Sovrascrivere file esistenti
  - Eliminare ZIP dopo l’estrazione
  - Cercare ZIP anche nelle sottocartelle
  - Mantenere la struttura delle sottocartelle  
- **Progress bar globale** per monitorare l’avanzamento  
- Log in tempo reale delle operazioni eseguite  
- Interfaccia utente semplice e intuitiva, con pulsanti per aprire la destinazione

---

## 💻 Requisiti

- Windows 10 o Windows 11  
- PowerShell 5.1 o superiore  
- Nessuna libreria esterna richiesta (usa solo **System.Windows.Forms** e **System.Drawing**)  

---

## 🚀 Installazione

1. Scarica o clona il repository GitHub:

```bash
git clone https://github.com/USERNAME/MultiZip.git
