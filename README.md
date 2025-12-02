# Pallinatore Quote v6

Applicazione per la numerazione automatica di quote su disegni tecnici.

## Funzionalità

- 📂 Apre immagini (PNG, JPG, TIF, BMP) e PDF
- 🔍 Scansione OCR automatica con PaddleOCR
- 🎯 Posizionamento automatico pallini numerati
- 🖱️ Pallini trascinabili con mouse
- 🔢 Rinumerazione automatica
- 📊 Esportazione in Excel
- 🖼️ Esportazione immagine pallinata
- 📄 Esportazione PDF

## Download

Vai alla sezione [Actions](../../actions) e scarica l'ultimo artifact "Pallinatore-Windows".

## Utilizzo

1. **Apri** un disegno tecnico (immagine o PDF)
2. **Scansiona OCR** per rilevare il testo
3. **Auto Pallina** per posizionare i pallini automaticamente
4. **Trascina** i pallini per posizionarli correttamente
5. **Click destro** per eliminare pallini in eccesso
6. **Rinumera** per riordinare gli ID
7. **Esporta** in Excel, immagine o PDF

## Controlli

| Azione | Comando |
|--------|---------|
| Sposta pallino | Trascina con mouse |
| Elimina pallino | Click destro |
| Aggiungi pallino | Click sinistro su area vuota |
| Elimina da tabella | Doppio click sulla riga |

## Compilazione manuale

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name "Pallinatore" pallinatore_v6.py
```

## Licenza

Uso libero.
