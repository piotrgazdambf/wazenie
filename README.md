# Ważenie Surowca – Generator i Badania (WSGiB)

Projekt Google Apps Script dla arkusza **(WSGiB) Ważenie Surowca - Generator i Badania**: karta ważenia (KW/KWG), PLS, stany surowcowe, wydanie skrzyń, raporty.

## Zawartość

- **KARTA_WAZENIA.js** – eksport karty ważenia, PLS, PS, menu Stany (Prześlij do Stanów), AKCJE SKRZYN
- **PDKW.js** – przekazywanie danych z WSG do KW/KWG
- **TABELEKW.js** – layout KW/KWG, tabele jakości, show/hide bloków
- **ETYKIETA.js**, **PDRS.js** – etykiety i raporty
- **appsscript.json** – manifest Apps Script

Deploy: **clasp + repo w jednym** (po jednorazowym logowaniu):

```bash
npm install
npx clasp login
npm run sync
```

- **`npm run push:gas`** – tylko wypchnij do Google Apps Script (clasp push)
- **`npm run pull:gas`** – tylko ściągnij z GAS (clasp pull)
- **`npm run sync`** – wypchnij do GAS, potem do GitHub (clasp push + git commit + git push)

Plik **`.clasp.json`** (scriptId projektu) jest w repo – projekt GAS i GitHub są podłączone pod ten sam skrypt.

---

## Zapis projektu na GitHub

**1. Zainstaluj Git**  
- Pobierz: https://git-scm.com/download/win  
- Zainstaluj (domyślne ustawienia wystarczą).

**2. Otwórz terminal w folderze projektu**  
- W Cursor: Terminal → New Terminal  
- Lub PowerShell / CMD i przejdź do folderu:
  ```bash
  cd "C:\Users\Admin\Desktop\Ważenie"
  ```

**3. Inicjalizacja repozytorium i pierwszy commit**
  ```bash
  git init
  git add .
  git commit -m "Pierwszy commit - projekt Ważenie Surowca"
  ```

**4. Utwórz repozytorium na GitHubie**  
- Wejdź na https://github.com/new  
- Nazwa np. `wazenie-surowca` (lub inna)  
- **Nie** zaznaczaj "Add a README" – repozytorium ma być puste.  
- Kliknij **Create repository**.

**5. Podłącz zdalne repozytorium i wypchnij**
  ```bash
  git remote add origin https://github.com/TWOJ_LOGIN/nazwa-repo.git
  git branch -M main
  git push -u origin main
  ```
  Zastąp `TWOJ_LOGIN` i `nazwa-repo` swoim kontem GitHub i nazwą repozytorium.

**Uwaga:** Po sklonowaniu repozytorium wykonaj `npm install` i (raz) `npx clasp login`. Potem `npm run sync` wypycha zmiany i do GAS, i na GitHub.
