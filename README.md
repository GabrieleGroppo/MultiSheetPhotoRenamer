# 📸 Multi Sheet Awesome Photo Renamer (MSAFR)  

MSAFR è uno script Python progettato per rinominare in batch le foto basandosi su dati provenienti da file Excel. Supporta più marchi e utilizza codici EAN per garantire un rinominamento preciso e automatizzato.  

## 🚀 Funzionalità  
✅ Estrazione automatica dei dati dai file Excel  
✅ Riconoscimento dei brand con colonne personalizzate  
✅ Ottimizzazione delle immagini JPEG per ridurre lo spazio  
✅ Generazione di report sui file rinominati e quelli mancanti  
✅ Supporto per diverse stagioni e strutture di cartelle  

## 🛠️ Requisiti  
- Python 3.x  
- pandas  
- openpyxl  
- jpegoptim (per l'ottimizzazione delle immagini)  

## 🔧 Installazione  
```bash
git clone https://github.com/GabrieleGroppoUni03/MSAFR.git
cd MSAFR
pip install -r requirements.txt
```
### requirements.txt
```plain-text
pandas
openpyxl
```

### jpegoptim
https://github.com/tjko/jpegoptim

## 📂 Utilizzo  
Esegui lo script passando la stagione e il brand:  
```bash
python MSAPRO.py stagione brand
```

Formato cartelle
```
📁 Progetto (dove tieni lo script .py)
└── 📁 pe25                 <-- Season name Folder
    ├── 📁 excels           <-- 'excels' folder
    │   └── liujo.xlsx      <-- Excel file must be named as brand
    └── 📁 photoes          <-- 'photos' folder
        └── 📁 liujo        <-- brand name subfolder
            ├── foto1.jpg   <-- Photos goes here
            └── foto2.jpg
```

## 📜 Licenza  
GNU License  

---



