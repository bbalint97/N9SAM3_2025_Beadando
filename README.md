# Beosztásból ICS generáló alkalmazás  

---

## Hallgató
- Név: **Berczi Bálint**
- Neptun-kód: **N9SAM3**

---

## Feladat leírása

A program célja egy Excel táblázatban tárolt (fiktív) műszakbeosztás alapján  
iCalendar (ICS) formátumú naptárfájl automatikus generálása.

A program létrejöttét valós probléma ihlette.




A program működése:

- Betölti a felhasználó által kiválasztott hónapnak megfelelő munkalapot.
- Megjeleníti a dolgozók listáját név és dolgozószám formátumban.
- A kiválasztott dolgozó műszakadataiból naptáreseményeket hoz létre.
- A cellaszínek alapján kezeli a készenlétet (piros cella).
- Figyelembe veszi a hónap napjainak számát (28/30/31 nap).
- Grafikus felületen indítható a generálást.
- A generált fájl ICS formátumban kerül mentésre.

---

## Modulok és a modulokban használt függvények

### `beosztas_BB_modul.py` 
Ez a fő programlogikát és a grafikus felületet tartalmazza.

### Saját függvények

| Függvény | Leírás |
|----------|--------|
| `bb_add_esemeny` | Egy ICS esemény összeállítása a megadott időpontokkal. |
| `bb_ics_generalas` | A teljes ICS fájl létrehozása a műszakadatokból. |
| `csomagellenorzes` | Ellenőrzi és szükség esetén telepíti a hiányzó modulokat. |
| `frissit_dolgozok` | Betölti a dolgozólistát a kiválasztott munkalap alapján. |
| `run` | Elindítja az ICS generálási folyamatot. |
| `log` | Üzeneteket ír a státuszmezőbe. |

### Modulok

**<font color=ORANGE>FONTOS: Az alkalmazás ellenőrzi a modulok létezését, hiány esetén autómatikusan telepíti</font>**

- `pandas`
- `openpyxl`
- `tkinter`
- `datetime`
- `calendar`

---

## Osztály

### `BB_AppGUI`
A modul saját osztálya, amely:

- elkészíti a grafikus felületet,
- létrehozza a legördülő listákat (munkalap, dolgozó),
- kezeli az eseményeket (gombnyomás, kiválasztás),
- elindítja a feldolgozási folyamatot,
- megjeleníti a státuszüzeneteket


---

## Grafikus felület

A program a `tkinter` és `ttk` modulokra épül.  
A GUI elemei:

- hónapválasztó legördülő lista,
- dolgozóválasztó legördülő lista,
- futtatás gomb,
- státusz szövegmező


---

## Eseménykezelés

A program több eseményt kezel:

- `<<ComboboxSelected>>` esemény a munkalap kiválasztásához,
- `command=self.run` a generáló gomb aktiválásához,
- státuszfrissítés a `log()` függvényen keresztül.

