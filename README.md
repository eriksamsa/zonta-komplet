# ZONTA KOMPLET – Global Avto

Windows/Python aplikacija za:

- iskanje rezervnih delov po znamki, modelu, letniku, motorju in nazivu dela,
- prikaz interne cene ter spletno preverjanje cene,
- evidenco motornih računalnikov ECU,
- OCR branje kataloških številk iz fotografij,
- shranjevanje ECU slik in iskanje po kataloški številki.

## Odpri v VS Code

1. Razpakiraj mapo.
2. V VS Code izberi **File > Open Folder...** in odpri mapo `ZONTA_KOMPLET_VSCODE`.
3. Odpri Terminal v VS Code.
4. Ustvari virtualno okolje:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

5. Zaženi:

```powershell
python src\main.py
```

Lahko tudi pritisneš **F5** v VS Code; priložen je `.vscode/launch.json`.

## OCR za ECU

Za dejansko branje kataloških številk iz slik mora biti na Windows nameščen **Tesseract OCR**. Program preverja običajni poti:

- `C:\Program Files\Tesseract-OCR\tesseract.exe`
- `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`

## Pomembne datoteke

- `src/main.py` – glavna aplikacija
- `requirements.txt` – Python knjižnice
- `scripts/run_windows.bat` – zagon na Windows
- `scripts/build_exe_windows.bat` – izdelava `.exe`
- `data/` – prostor za bazo in slike
- `.vscode/` – nastavitve za VS Code

## Opomba o podatkih

Aplikacija trenutno ustvarja `ECU_BAZA.xlsx` in mapo `ECU_SLIKE` poleg izvedbene Python datoteke. Če želiš, lahko to v naslednji verziji prestavimo v mapo `data/` in dodamo samodejno shranjevanje nastavitev ter poti do glavne evidence.
