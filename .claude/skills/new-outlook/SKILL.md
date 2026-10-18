---
name: new-outlook
description: Automatizacija "New Outlook" (olk.exe, WebView) na Windowsu bez COM-a — preko UI Automation (pywinauto) i simuliranih tastera (pyautogui). Koristi kad treba snimiti mejl kao .eml, premestiti ga u folder, prepoznati selektovani mejl ili Teams pozivnicu, ili kad win32com/Outlook.Application ne radi.
---

# New Outlook bez COM-a

New Outlook (`olk.exe`) je WebView aplikacija i **ne podrzava COM** —
`win32com.client.Dispatch("Outlook.Application")` ne vidi njegove mejlove.
Ostaju dva puta:

1. **Microsoft Graph API** — zvanican i najpouzdaniji (`/me/messages/{id}/$value`
   daje MIME = .eml, `/move` premesta). Trazi Azure app registraciju i odobrenje IT-a.
   Ako je to moguce, koristi ovo umesto svega ispod.
2. **UI automatizacija** — UIA (`pywinauto`, backend `uia`) za *citanje* stanja,
   `pyautogui` za tastere, `pyperclip` za tekst. Ovaj dokument opisuje taj put.
   Referentna implementacija: `save_emails.py` u ovom repou.

Zavisnosti: `pyautogui pyperclip keyboard pywinauto`.

> Nazivi UIA elemenata ("About Outlook", "Meeting") vaze za **englesko** korisnicko
> okruzenje. Precice su proverene na New Outlooku iz 2026. Posle azuriranja
> Outlooka ponovo proveri recepte (`dry_run = true` u `config.ini`).

---

## Pronalazenje prozora

```python
from pywinauto import Desktop
win = next((w for w in Desktop(backend="uia").windows()
            if "Outlook Host" in (w.class_name() or "")), None)
```

Prozor ume da nestane/kasni pri startu — pokusaj vise puta sa pauzom.

## Fokus na prvi mejl u listi

```
Esc → Ctrl+1 → Esc → Alt → F6 → F6 → F6 → Home → Down → Up
```

(`win.set_focus()` pre toga; ~0.05–0.25 s izmedju tastera.)

## Koji item je selektovan

Trazi `ListItem` sa keyboard fokusom; njegov `name` sadrzi posiljaoca, predmet i vreme.

```python
for elem in win.descendants():
    ei = elem.element_info
    if ei.control_type == "ListItem" and ei._element.CurrentHasKeyboardFocus:
        name = ei.name   # "Meeting ..." = Teams/meeting pozivnica
```

`descendants()` je spor (sekunde) — ne zovi ga u petlji bez potrebe.

## Snimanje kao .eml

**Obican mejl:** `Alt → F` (File meni). Proveri da se meni otvorio tako sto u UIA
stablu postoji element sa imenom **"About Outlook"**. Ako ga nema, fokus nije na
mejlu (npr. group header) — Esc nekoliko puta, sacekaj, ponovi.
Ako ga ima: `A → Enter` otvara Save As.

**Teams/meeting pozivnica** (File meni ne radi): kontekst meni
`Shift+F10 → End → Up×3 → Right → Enter`.

**Save As dialog** je klasican Win32 dialog (`#32770`, backend `win32`), prepoznaje
se po deci `"&Save"` ili `"Email Message (*.eml)"`. Predlozeno ime fajla = predmet.
U polje za naziv upisi **punu putanju** — dialog sam prelazi u folder:

```python
pyautogui.hotkey("alt", "n"); pyautogui.hotkey("ctrl", "a")
pyperclip.copy(full_path); pyautogui.hotkey("ctrl", "v"); pyautogui.press("enter")
```

Posle Enter **cekaj da fajl postoji na disku** (ne veruj pauzi).

### Kada nema Save As dialoga → Downloads

New Outlook ponekad (npr. veliki prilozi, mejl nije lokalno ucitan) umesto
dialoga **preuzme .eml u Downloads**. Postupak:

1. Snimi sadrzaj Downloads foldera **pre** akcije.
2. Cekaj da se pojavi bilo koji nov fajl (download poceo), pa da postane `.eml`.
3. Procitaj `Subject` iz zaglavlja (`email.message_from_binary_file` +
   `email.header.decode_header`), preimenuj i premesti gde treba.

### Lazni error dialozi

- **"File not found" / naslov "Warning: This site can see edits you make"** —
  WebView bug. Zatvori (Enter), zatvori Save As (Esc), sacekaj ~10 s, ponovi.
- **"Confirm Save As"** (fajl postoji) — Esc i izaberi drugo ime (`_1`, `_2`...).

Prepoznavanje: foreground prozor (`GetForegroundWindow`) klase `#32770` koji nije
postojao pre Enter-a; tekst iz naslova i `EnumChildWindows`.

## Premestanje u folder

```
Ctrl+Shift+V → (paste naziv foldera) → Enter
```

Folder mora postojati; naziv se lepi iz clipboarda.

---

## Pravila pouzdanosti

- **Pre destruktivne akcije (premestanje/brisanje) proveri da je selektovan isti
  mejl** — nov mejl moze stici na vrh liste izmedju dva koraka. Poredi predmet
  (samo slova i cifre, lowercase) sa nazivom fokusiranog `ListItem`-a.
- Tekst uvek **lepi** (pyperclip), ne kucaj — raspored tastature i nasa slova.
- Sva kasnjenja u config; spori racunari traze vece.
- Ukljuci `pyautogui.FAILSAFE = True` i hotkey za hitno gasenje (`os._exit`) —
  skripta salje tastere u prozor koji ima fokus, kakav god da je.
- Svaki prekid batch-a mora vratiti stanje (`try/finally`), inace hotkey ostane
  "zaglavljen".
- Loguj svaki korak (CSV + kopija konzole u fajl).
- Korisnik ne sme da koristi racunar dok batch radi.
