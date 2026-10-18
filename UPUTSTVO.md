# Uputstvo za rad — Outlook EML Saver

## Sta radi

Skripta automatski snima selektovani mejl iz Outlooka kao .eml fajl na disk
i premesta ga u odredisni folder (npr. Inbox2). Koristi se iskljucivo tastaturom.

---

## Pokretanje

```
python save_emails.py
```

Pri startu skripta pita u koji Outlook folder da premesta mejlove:

```
1. Inbox2  (default)
2. Sent2
3. Archive2
4. Servis obradjeno2
5. Unesi drugi naziv
```

Pritisni Enter za default (Inbox2) ili unesi broj/naziv. Skripta zatim pocinje da slusa.

---

## Hotkey

| Tastatura       | Akcija                                               |
|-----------------|------------------------------------------------------|
| Ctrl+Shift+S    | Snimi selektovani mejl i premesti u odredisni folder |
| Ctrl+Shift+S    | Ako je batch aktivan — zaustavi posle tekuceg mejla  |
| Enter (u CMD)   | Zatvori skriptu                                      |

---

## Tok rada

1. Otvori Outlook, klikni na mejl koji hoces da snimis
2. Pritisni **Ctrl+Shift+S**
3. Skripta automatski:
   - Vraca fokus na listu mejlova
   - Otvara kontekst meni (desni klik) i proverava da li je mejl selektovan
   - Navigira do Save As → Email Message (*.eml)
   - Snima fajl u SAVE_FOLDER sa nazivom: `DATUM_Predmet.eml`
   - Premesta mejl u odredisni Outlook folder
   - Automatski prelazi na sledeci mejl i ponavlja

---

## Fajlovi koje skripta kreira

| Fajl            | Sadrzaj                                        |
|-----------------|------------------------------------------------|
| `log.csv`       | Evidencija svakog snimljenog mejla             |
| `terminal.log`  | Kompletan ispis iz CMD prozora                 |

---

## Kako radi (tehnicko)

**Detekcija mejla:**
Skripta otvara kontekst meni sa Shift+F10 i pretrazuje UIA stablo Outlook prozora
trazeci element "Save as". Ako ga nema, znaci da je selektovan grupni header (datum/naziv
grupe), a ne mejl — skripta ceka i pokusava ponovo.

**Snimanje:**
Navigira kroz meni: End → 3x Up → desno → Enter, sto otvara Save As dialog.
Dialog se popunjava automatski (naziv fajla, folder) i potvrduje sa Enter.
Skripta ceka da se fajl fizicki pojavi na disku pre nego sto nastavi.

**Velika prilozena fajlovi:**
Ako mejl ima veliku prilogu, Outlook ga preuzima sa servera pre snimanja.
Skripta detektuje preuzimanje i ceka da se .eml fajl pojavi u Downloads folderu,
zatim ga preimenuje i premesta u SAVE_FOLDER.

**Premestanje:**
Posle snimanja, skripta ponovo otvara kontekst meni (Shift+F10) i navigira
do "Move to folder", ukucava naziv odredisnog foldera i potvrdjuje.

**Batch rezim:**
Skripta moze da obradi do 100 mejlova bez prekida. Posle svakog mejla automatski
prelazi na sledeci. Pritisak Ctrl+Shift+S u toku batch-a zaustavlja obradu
posle trenutnog mejla.

**Greske:**
- Ako se pojavi Windows "File not found" dialog (Outlook bug), skripta ga zatvara
  i ponovo pokusava snimanje
- Ako mejl nije dostupan (nije ucitan), ceka i pokusava do 5 puta
- Ako Outlook nije pronadjen, pokusava do 5 puta sa pauzom od 10 sekundi
