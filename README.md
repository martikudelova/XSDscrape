# XSD → XLSX Automation

Tento projekt slouží k automatickému zpracování XML Schema (`.xsd`) souborů a vytvoření přehledného Excel výstupu (`.xlsx`) se strukturou dokumentu.

Skript projde celé XSD, najde všechny **leaf elementy** a ke každému určí:

- Hierarchickou cestu
- ISO status (M / O / C)
- Typ elementu
- Odvozený formát (např. `X(35)`, `XN(15)`, `ISODate`, `boolean`, `<ANY>`)

---

## Požadavky

- Python 3.10 nebo novější  
- Operační systém: Windows  
- Přístup k příkazovému řádku (cmd.exe)

---

## User Guide

1. Vytvoř si lokální adresář, např.:  
   `C:\Users\<uživatel>\Documents\XSD_to_XLSX`

2. Přesuň soubor `XSDscrape.py` a `.xsd` soubor, který chceš zpracovat, do vytvořeného adresáře.

3. Spusť příkazový řádek (cmd.exe).

4. Přejdi do nového adresáře:  
   cd C:\Users\<uživatel>\Documents\XSD_to_XLSX

5. Vytvoř si virtuální prostředí (pouze jednou):  
   python -m venv venv

6. Aktivuj virtuální prostředí:  
   .\venv\Scripts\Activate

7. Nainstaluj potřebnou knihovnu (pouze jednou):  
   pip install openpyxl

8. Spusť skript:  
   python XSDscrape.py

9. Zadej název `.xsd` souboru, který chceš zpracovat.

10. Výstupní `.xlsx` soubor najdeš ve stejném lokálním adresáři.

---

## Output

Excel obsahuje **dva listy**:

### 1️⃣ Hierarchy – hierarchie všech leaf elementů

Sloupce:

- `Level 1…N` – jednotlivé úrovně
- `Full Path` – kompletní cesta elementu
- `Type name` – název typu
- `ISO Status` – M (mandatory), O (optional), C (conditional)
- `Format` – odvozený formát
- `Patterns` – pattern
- `Enumerations` – seznam enumerací

### 2️⃣ Types – seznam všech unikátních typů a jejich formátů

Sloupce:

- `type` – název typu
- `format` – odvozený formát
- `minLength` – minimální délka
- `maxLength` – maximální délka
- `totalDigits` – počet číslic
- `fractionDigits` – počet desetinných míst
- `pattern` – regex pattern
- `enumeration` – seznam enumerací
