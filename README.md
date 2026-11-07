# Radar Adatok Vizualizátora

Windowsos asztali alkalmazás radaradatok megjelenítésére, összehasonlítására és
összesítésére. A program Python és Tkinter segítségével készült; a grafikonokat
a Matplotlib rajzolja.

## Funkciók

- ZDR- és Zátlag-adatok megjelenítése magasság szerint.
- Két mérési fájl összehasonlítása, valamint a különbségek elemzése.
- Havi ZDR- és magasság-pontfelhő, napi átlagolási és statisztikai lehetőségekkel.
- Data Quality Control (DQC) szűrés 12 oszlopos adatokra.
- Kalibrálás előtti és utáni ZDR-adatok összesítése Excel-fájlba.
- Olvadásiréteg-adatok havi átlagolása.
- Ábrák mentése képfájlba; az olvadásiréteg-ábra PNG, PDF, SVG vagy JPEG formátumban is menthető.
- Opcionális radarkép-betöltés a fájlnévben szereplő mérési idő alapján.

## Követelmények

- Windows 10 vagy újabb.
- Python 3 telepített Tkinter támogatással. A Windowsos, python.org-ról letöltött
  Python-telepítő általában tartalmazza a Tkintert.
- Internetkapcsolat csak a radarképek online betöltéséhez szükséges.

## Telepítés

1. Töltsd le vagy másold a projekt fájljait egy helyi mappába.
2. Telepítsd a Pythont a [python.org](https://www.python.org/downloads/windows/)
   oldalról. A telepítőben jelöld be az **Add Python to PATH** lehetőséget.
3. Nyiss egy PowerShell-ablakot a projekt mappájában, és hozd létre a virtuális
   környezetet:

   ```powershell
   py -m venv .venv
   ```

4. Telepítsd a szükséges csomagokat:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

   A függőségek a [requirements.txt](requirements.txt) fájlban vannak felsorolva.

## Indítás

A telepítés után kétféleképpen indíthatod:

- Kattints duplán az `Inditas.bat` fájlra.
- PowerShellből futtasd:

  ```powershell
  .\.venv\Scripts\python.exe .\radar_plotter.py
  ```

A parancsikon létrehozásához futtasd egyszer a `Parancsikon_letrehozasa.vbs`
fájlt. Ez az Asztalon létrehoz egy **Radar Plotter** parancsikont, amely az
`Inditas.bat` fájlt indítja.

## Bemeneti fájlok

Az alkalmazás `.txt` fájlokból olvas. Az adatokat soronként, szóközzel vagy
tabulátorral elválasztva add meg; minden adatsor egy magassághoz tartozik.

### Alapformátum: 3 oszlop

Az első oszlop a magasság km-ben, a második a ZDR, a harmadik a Zátlag:

```text
0.5 1.2 15.3
1.0 -0.5 12.8
1.5 NAN NAN
```

Hiányzó értéket `NAN` jelölhet. A hibás vagy nem numerikus sorokat a program
kihagyja.

### Bővített formátum: 12 oszlop

Az első oszlop a magasság, az utána következő 11 oszlop sorrendje:

```text
Magasság ZDR ZH ZV UH UV P RHOHV SNRV SNRH CCORH CCORV
```

Például:

```text
0.5 1.2 15.3 14.8 2.1 2.0 0.1 0.99 32.0 31.5 0.0 0.0
```

A 12 oszlopos adatokhoz tartozó kiegészítő grafikonokon és a DQC-fülön ez a
formátum használható. A DQC-feldolgozás olyan mappát vár, amely közvetlenül
tartalmaz `.txt` fájlokat, vagy radarnevű almappákat, bennük `.txt` fájlokkal.

### Fájlnév dátummal és állomáskóddal

A dátumhoz kötött funkciókhoz a fájlnév elején az állomáskódot és az időpontot
használd ebben a formában:

```text
BUD202509121430.txt
```

Vagyis: `ÁLLOMÁSKÓDÉÉÉÉHHNNÓÓPP.txt`. Felismert állomáskódok: `BUD`, `HAR`,
`NAP`, `POG`, `SZE`.

A hónap szerinti pontfelhő egy kiválasztott hónapmappán belül napokra bontott
almappákat keres, azokban dátumot tartalmazó nevű `.txt` fájlokkal. Az
olvadásiréteg-fül egy év mappát vár, amely hónapnevekkel vagy hónapszámokkal
elnevezett almappákat, illetve közvetlenül hónapot jelölő nevű `.txt`
fájlokat tartalmazhat.

## Használat

1. Az ablak felső részén kattints a **Fájl megnyitása (.txt)** gombra, és válaszd
   ki az első mérési fájlt.
2. Az adatok megjelennek a magasság–ZDR grafikonon; a füleken további
   összehasonlító, összesítő és szűrőeszközök érhetők el.
3. A grafikonok nagyításához, mozgatásához és képként mentéséhez használd a
   grafikon alatti eszköztárat, ahol az elérhető.
4. Az **Összehasonlítás** fülön tölts be egy második fájlt az elsővel való
   összevetéshez.
5. A mappaalapú füleken válaszd ki a megfelelő hónap-, év- vagy adatgyűjtő
   mappát. Az Excel-exporthoz válaszd ki a kalibrálás előtti és utáni mappát,
   majd kattints az exportálás gombra.

Ha egy fület vagy műveletet nem használsz, nem szükséges hozzá adatot betöltened.
Az ablak alján megjelenő állapotüzenetek jelzik a feldolgozás eredményét és a
hibákat.

## Excel-sablon

Az Excel-exporthoz a projektben található `tablazat.xlsx` használható
sablonként. Az exportálás során válaszd ki a kalibrálás előtti és utáni
adatmappát; igény szerint másik `.xlsx` sablont is megadhatsz, majd válaszd ki
az új fájl mentési helyét. Az eredmény ZDR-átlagokat és a kalibrálás
előtti-utáni változásokat tartalmazza.

## Adatvédelem

A betöltött helyi mérési fájlokat a program helyben olvassa és dolgozza fel; az
adatok feltöltése nem része a program működésének. Ha a fájlnév alapján
radarkép kérhető le, a program a dátumhoz tartozó nyilvános MetNet-radarképet
tölti le. Ehhez internetkapcsolat kell. Nem nyilvános mérési adatokat csak
jogosultság esetén használj, és megosztás előtt ellenőrizd az exportált
fájlokat is.

## Hibaelhárítás

- **A Python-parancs nem található:** telepítsd újra a Pythont, és engedélyezd
  az **Add Python to PATH** beállítást; ezután nyiss új PowerShell-ablakot.
- **Nem indul az `Inditas.bat`:** ellenőrizd, hogy a projektmappában létezik-e
  a `.venv` mappa. Ha hiányzik vagy sérült, ismételd meg a telepítési lépéseket.
- **Hiányzó modulra panaszkodik:** a projektmappában futtasd újra a
  `.\.venv\Scripts\python.exe -m pip install -r requirements.txt` parancsot.
- **A Tkinter hiányzik:** módosítsd vagy telepítsd újra a Python-telepítést úgy,
  hogy a Tcl/Tk komponens is települjön.
- **Nem jelenik meg adat vagy radarkép:** ellenőrizd a fájl oszlopszámát és
  formátumát, a dátumos fájlnév szerkezetét, illetve radarkép esetén az
  internetkapcsolatot.
