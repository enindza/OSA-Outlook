# Uputstvo za rad — Outlook EML Saver

## Sta radi

Skripta automatski snima mejl sa vrha liste u Outlooku kao .eml fajl na disk
i premesta ga u odredisni folder (npr. Inbox2). Radi u batch-u — posle svakog
mejla prelazi na sledeci. Koristi se iskljucivo tastaturom.

---

## Instalacija (jednom)

```
pip install -r requirements.txt
```

Kopiraj `config.ini.example` u `config.ini` (isti folder kao skripta) i podesi
putanje (`save_folder`, `downloads_folder`).

---

## Pokretanje

Dupli klik na `POKRENI.bat`, ili:

```
python save_emails.py
```

Pri startu skripta pita u koji Outlook folder da premesta mejlove
(lista se podesava u `config.ini` → `dest_folder_options`):

```
1. Inbox2  (default)
2. Sent2
3. Archive2
4. Servis obradjeno2
5. Unesi drugi naziv
```

Pritisni Enter za default ili unesi broj/naziv. Skripta zatim pocinje da slusa.

---

## Hotkey

| Tastatura                    | Akcija                                                  |
|------------------------------|---------------------------------------------------------|
| Ctrl+Shift+S                 | Pokreni batch: snimi mejl sa vrha i premesti ga         |
| Ctrl+Shift+S (tokom batch-a) | Zaustavi posle tekuceg mejla                            |
| Ctrl+Shift+X                 | Hitno gasenje skripte — odmah                           |
| Mis u gornji levi ugao       | Hitno zaustavljanje batch-a (pyautogui failsafe)        |
| Enter (u CMD)                | Zatvori skriptu                                         |

Hotkey-evi se menjaju u `config.ini` → `[hotkeys]`.

---

## Tok rada

1. Otvori Outlook
2. Pritisni **Ctrl+Shift+S**
3. Skripta za svaki mejl automatski:
   - Vraca fokus na gornji mejl u listi
   - Prepoznaje da li je obican mejl ili Teams/meeting pozivnica
   - Otvara Save As i snima fajl u `save_folder` sa nazivom `DATUM_Predmet.eml`
   - Proverava da je na vrhu liste i dalje isti mejl
   - Premesta mejl u odredisni Outlook folder
   - Prelazi na sledeci mejl (do `max_batch` mejlova)

Dok batch radi, ne koristi racunar — skripta salje tastere u prozor koji ima fokus.

---

## Fajlovi koje skripta kreira

| Fajl            | Sadrzaj                                        |
|-----------------|------------------------------------------------|
| `log.csv`       | Evidencija svakog snimljenog/premestenog mejla |
| `terminal.log`  | Kompletan ispis iz CMD prozora                 |

Statusi u `log.csv`: `ok`, `moved`, `move_skipped`, `no_dialog`, `skip`,
`error`, `aborted`, `batch_done`, `batch_stopped`.

---

## Kako radi (tehnicko)

**Fokus na mejl:**
Esc → Ctrl+1 (Mail) → Esc → Alt → 3×F6 → Home, sto stavlja fokus na prvi mejl u listi.

**Obican mejl:**
Alt → F otvara File meni. Skripta u UIA stablu trazi element "About Outlook" —
ako ga nema, File meni se nije otvorio (fokus nije na mejlu, npr. na group
header-u), pa ceka 5s i pokusava ponovo (do `max_group_retry` puta).
Ako ga ima: A → Enter otvara Save As dialog.

**Teams/meeting pozivnica:**
Prepoznaje se po UIA nazivu itema (pocinje sa "Meeting"). Save As se otvara
preko kontekst menija: Shift+F10 → End → 3×Up → Right → Enter.

**Snimanje:**
U Save As dialogu se u polje za naziv upisuje puna putanja
(`save_folder\DATUM_Predmet.eml`) i potvrdjuje sa Enter. Ako fajl sa tim
imenom vec postoji, dodaje se `_1`, `_2`, ... Skripta ceka da se fajl fizicki
pojavi na disku pre nego sto nastavi.

**Bez Save As dialoga (Downloads):**
Ako se Save As dialog ne pojavi, Outlook je mejl verovatno preuzeo u Downloads.
Skripta ceka da se nov .eml pojavi u `downloads_folder`, cita mu Subject,
preimenuje ga i premesta u `save_folder`.

**Premestanje:**
Posle snimanja skripta ponovo stavlja fokus na gornji mejl i proverava da je
to isti mejl koji je snimljen (predmet mora biti sadrzan u nazivu itema).
Ako je u medjuvremenu stigao nov mejl, NE premesta ga — upisuje
`move_skipped` u log i prekida batch. Ako je isti: Ctrl+Shift+V → naziv
foldera → Enter. Provera se iskljucuje sa `verify_before_move = false`.

**Greske:**
- Windows "File not found" dialog (Outlook bug) ili "Confirm Save As" —
  skripta ga zatvara i ponavlja snimanje
- Nema fajla ni u Downloads — ponavlja do `max_dialog_retry` puta,
  sa pauzom `dialog_retry_wait` sekundi
- Outlook nije pronadjen — pokusava do 5 puta sa pauzom od 10 sekundi
- Svaka greska koja prekine batch vraca hotkey u normalno stanje, pa sledeci
  Ctrl+Shift+S ponovo pokrece batch
