# Motor „Domaćin radnje" bez ličnog računara: predlog arhitekture i plan prelaska

Stanje: predlog, oktobar 2026. Osnova je opis postojećeg motora (PowerShell + klasični Outlook COM + Windows Task
Scheduler). Pretpostavke su označene sa **[P]**. Stavke koje se moraju potvrditi u probi označene su sa **[proveriti]**.

---

## 0. Ukratko

- **Preporuka:** Power Automate cloud flow-ovi na **standardnim M365 konektorima**. Pokreće ih **licencirani servisni
  nalog**. Ulazna i izlazna pošta ide preko **zajedničkog sandučeta `domacin@`**, a stanje motora se čuva u
  **SharePoint listama** na postojećem sajtu.
- Složenija logika (čišćenje teksta odgovora, rokovi u radnim satima, HTML izveštaja, grupisanje podsetnika) prelazi u
  **Office Scripts**. To je TypeScript nad jednom Excel radnom sveskom na sajtu, a pokreće se standardnom akcijom
  *Run script*.
- **Jedini deo za koji standard možda nije dovoljan je PDF sa ugrađenim fotografijama.** Prvo se proba besplatna
  konverzija HTML → PDF preko OneDrive-a. Ako kvalitet nije dobar, uvodi se mala **Azure Function** (headless Chromium).
  Za nju treba Premium (HTTP) za servisni nalog i Azure pretplata.
- **Dugmad ostaju `mailto`**, ali na `domacin@` umesto lične adrese. Approvals i Adaptive Cards se ne preporučuju za
  ovaj tok (obrazloženje je u odeljku 4).
- Tvoj računar posle prelaska služi samo za analitiku (DuckDB), kao i do sada.

---

## 1. Arhitektura

```
 Domaćin (PWA)                         Službe / odobravaoci (samo mejl)
   │  fajl posete + slike                    │  odgovor u niti, mailto dugmad
   ▼                                         ▼
 SharePoint  posete/  ◄── prilozi ───  Zajedničko sanduče domacin@  (Exchange Online)
   │                                         │
   │ F2 Obrada poseta (svaki sat 10–17)      │ F1 Prijem pošte (na svaku poruku)
   ▼                                         ▼
 ┌──────────────────────── SharePoint /sites/domacin ────────────────────────┐
 │ DomacinZahtevi (postojeća)   Motor-Posete   Motor-Odobrenja   Motor-Poruke │
 │ Motor-IzlazniMejlovi   Motor-Log   Motor-Podesavanja   Motor-PravilaStavki │
 │ Motor-Praznici   Motor-UgovoriFasada   izvestaji/<period>/*.pdf            │
 └───────────────────────────────────────────────────────────────────────────┘
   ▲            │                                   ▲
   │ F4 Rokovi i podsetnici (svaki sat)             │ F5 PDF izveštaja
   │            ▼                                   │
   │   F0 Slanje mejla (iz liste Motor-IzlazniMejlovi, prekidač PROBA) ──► domacin@ šalje
   │
 Office Scripts (Motor-logika.xlsx): ocistiOdgovor, rokRadnihSati, izvestajHtml, podsetniciGrupisi
```

### 1.1 Šta gde ide

| Deo | Gde | Zašto |
|---|---|---|
| Okidači, raspored, čitanje i slanje pošte | Power Automate cloud flow | Radi u oblaku. Konektori Office 365 Outlook i SharePoint su standardni. |
| Upis u listu, prilozi stavki | SharePoint konektor (*Create item*, *Update item*, *Add attachment*) | Standardni konektor, radi pod servisnim nalogom. |
| **Komentari liste i @oznaka** | SharePoint akcija **Send an HTTP request to SharePoint** → `_api/web/lists/GetByTitle('…')/items(<id>)/Comments` | Standardna akcija, iako je REST (to nije Premium HTTP). Rešava rupu koju Graph ima. **[proveriti]** format `mentions` u telu zahteva. |
| Stanje umesto lokalnog CSV/JSON-a | Nove SharePoint liste (1.2) | Trajno, vidljivo, sa verzijama. Nije vezano za računar. |
| Logika koju je teško pisati u izrazima | Office Scripts (*Excel Online (Business) › Run script*) | Standardno. Kod je sličan postojećem, a lakše se održava nego stotine izraza. |
| PDF | 1. OneDrive *Convert file* (HTML → PDF), 2. rezerva: Azure Function sa Chromium-om | Odeljak 5. |
| Analitika | Ostaje lokalna | Kao do sada. Čita listu i fajlove. |

**Entra aplikacija (`Sites.Selected`) motoru više nije potrebna.** Flow-ovi rade delegirano, kao servisni nalog. Aplikaciju
i sertifikat zadrži samo za lokalnu analitiku. Autor u listi biće „Domacin radnje (servis)" umesto „SharePoint App".

### 1.2 Nove liste za stanje (umesto lokalnih fajlova)

Važno: motor **ne sme da upisuje svoje stanje u redove `DomacinZahtevi`**. Svaki takav upis menja `Modified`, a
eskalacija posle 14 dana se računa baš po `Modified`. Zato stanje ide u posebne liste vezane za zahtev preko `ZahtevId`.

| Lista | Ključ (jedinstvena vrednost) | Kolone | Zamenjuje |
|---|---|---|---|
| **Motor-Posete** | `IdPosete` | Fajl, Sifra, Datum, VremeFajla, Status (Primljeno / Obrađeno / Greška), BrojZahteva, IzvestajPoslat, PdfPutanja, Greska | registar obrađenih poseta |
| **Motor-Odobrenja** | `ZahtevId` | Nivo (1/2), Korak (OPSTI / RUKOVODILAC / ZAMENIK / DIREKTOR), NaRedu (adrese), Rok, PodsetnikPoslat, Hitno, Procena, IstorijaJSON | lokalni JSON odobrenja |
| **Motor-Eskalacije** | `ZahtevId` | ModifiedPriEskalaciji, DatumEskalacije | „šta je eskalirano" |
| **Motor-Poruke** | `InternetMessageId` | Od, Naslov, Tip (TIKET / KOMANDA / POSETA / OSTALO), ZahtevId, Rezultat | kategorija „obrađeno" u Outlook-u |
| **Motor-IzlazniMejlovi** | (ID) | Tip, To, Cc, Naslov, TeloHtml (višeredni tekst), Prilozi (putanje, JSON), Status, Poslato, Greska | direktno slanje iz koda |
| **Motor-Podesavanja** | `Kljuc` | Vrednost. Ključevi: `REZIM` (PROBA / ZVANICNO), `PROBA_PRIMALAC`, `IZVESTAJ_TO`, `PRAG`, `NIVO2_PRAG`, `RUKOVODILAC`, `ZAMENIK`, `OPER_DIREKTOR`, `DL_OPSTI_POSLOVI`, `POSLEDNJI_PONEDELJAK` | konstante u kodu, „kada je poslat ponedeljni podsetnik" |
| **Motor-PravilaStavki** | `StavkaId` | Pravilo (BEZ / BEZBEDNOST / UVEK / PRAG), Fasada (da/ne) | tabela 90 pravila |
| **Motor-Praznici** | `Datum` | Naziv | (novo) praznici u rokovima |
| **Motor-UgovoriFasada** | `Sifra` | KoOdrzava, IzvodIzUgovora | lokalna baza ugovora o zakupu **[P]**: izvod je mali i retko se menja, pa se jednom izveze |
| **Motor-Log** | (ID) | Flow, Nivo, Poruka, Kontekst, RunUrl | lokalni log |

**Jedinstvena vrednost** (*Enforce unique values*) na ključnim kolonama je osnovna zaštita od duple obrade. Ako dva
pokretanja pokušaju da upišu istu poruku ili istu posetu, drugo dobija grešku i preskače posao.

---

## 2. Spisak flow-ova

Svi flow-ovi su u jednom **rešenju (solution)** u okruženju sa Dataverse-om. To omogućava child flow-ove, izvoz rezervne
kopije i prenos. **[P]** Ako IT ne odobri rešenje, sve radi i bez njega: zajednički deo ide preko liste
`Motor-IzlazniMejlovi` (F0), a child flow-ovi nisu potrebni.

| # | Naziv | Okidač | Glavni koraci | Konektori |
|---|---|---|---|---|
| **F0** | Slanje mejla | SharePoint: *When an item is created* u `Motor-IzlazniMejlovi` | Pročita `REZIM`. U PROBI: To = `PROBA_PRIMALAC`, Cc prazan, a na vrh tela ide crvena traka „U zvaničnom režimu ide: To … / Cc …". Učita priloge po putanjama (*Get file content using path*). *Send an email from a shared mailbox (V2)* kao `domacin@`. Upiše Status / Poslato / Greska. | SharePoint, Office 365 Outlook (standard) |
| **F1** | Prijem pošte | Outlook: *When a new email arrives in a shared mailbox (V2)*, `domacin@`, sa prilozima. **Concurrency = 1.** | 1) Upis u `Motor-Poruke` po `InternetMessageId`. Ako već postoji, kraj. 2) Ako je pošiljalac sam `domacin@` (motor), kraj. 3) Klasifikacija po naslovu: `\[Z-(\d+)\]` + ključna reč. **KOMANDA** (ODOBRENO / ODBIJENO / DOPUNA / DOPUNJENO / PROCENA) → grana odobravanja (2.1). **TIKET** → grana tiketa (2.2). Prilozi `poseta_*.txt/.jpg` → snimanje u `posete` (*Create file*, bez gaženja: prvo *Get file metadata using path*). Ostalo → samo log. 4) Premeštanje poruke u folder `Obrađeno`. | Outlook, SharePoint, Excel Online (Run script) — standard |
| **F2** | Obrada poseta | *Recurrence*: radnim danima, svakog sata 10:00–17:00, zona Central Europe | Lista `.txt` fajlova u `posete` izmenjenih od poslednjeg rada (*Get files (properties only)* sa filterom). Za svaki: parsiranje JSON-a, pa provera da su sve slike iz `stavke[].fotografije` stigle (ako nisu, preskače se do sledećeg sata, posle 3 sata šalje se upozorenje). Idempotentnost po `id`: noviji fajl istog `id` → ponovna obrada samo izveštaja, uz dopunu zahteva koji nedostaju. Upis u `Motor-Posete`. Za svaku stavku NE i svaki CM sektor: provera da li postoji zahtev (IdPosete + StavkaId + Kategorija), ako ne, *Create item* + *Add attachment* za slike. Pa mejl službi u `Motor-IzlazniMejlovi`. Zvanični izveštaj (`izvestajHtml`) → `Motor-IzlazniMejlovi`. Poziv F5. | SharePoint, Excel Online — standard |
| **F4** | Rokovi i podsetnici | *Recurrence*: radnim danima, svakog sata 10:00–17:00 | a) **Procene upisane u listi:** zahtevi sa `ProcenaTroska` a bez reda u `Motor-Odobrenja` → pokretanje odobravanja po pravilu. b) **Rokovi odobravanja:** redovi gde je `Rok < sada` → sledeći korak ili podsetnik nivoa 2. c) **Ponedeljak:** ako je danas ponedeljak i `POSLEDNJI_PONEDELJAK ≠ danas` → `podsetniciGrupisi` (jedan mejl po adresi, bez zahteva koji čekaju odobrenje), pa upis datuma. d) **Eskalacija 14 dana:** jednom dnevno, u prvom radu → po RM-u, preskače zahteve gde je `Modified = ModifiedPriEskalaciji`. | SharePoint, Excel Online — standard |
| **F5** | PDF izveštaja | Child flow iz F2 (ili *When an item is modified* u `Motor-Posete`, `IzvestajPoslat = da`) | `izvestajHtml` (verzija za štampu, slike kao data URI). Pa privremeni `.html` u OneDrive servisnog naloga → *Convert file* (PDF) → *Create file* u `izvestaji/<period>/izvestaj_<šifra>_<datum>_<HHMM>.pdf` (isto ime = nova verzija), pa brisanje privremenog fajla. Rezerva: Azure Function (odeljak 5). | OneDrive for Business, SharePoint — standard. Rezerva: HTTP — **Premium** |
| **F6** | Nadzor | *Recurrence*: radnim danima u 17:15 | Pregled grešaka iz `Motor-Log`, neposlatih redova u `Motor-IzlazniMejlovi` i poseta u statusu Greška. Ako ih ima, jedan mejl tebi. Uz to podešavanje **flow failure** obaveštenja za vlasnike. | SharePoint, Outlook — standard |

### 2.1 Grana odobravanja (u F1, concurrency 1)

1. Učitavanje zahteva i reda u `Motor-Odobrenja`.
2. `PROCENA <iznos>`: upis `ProcenaTroska` u zahtev. Pravilo stavke iz `Motor-PravilaStavki` određuje šta dalje:
   - BEZ: ništa;
   - BEZBEDNOST: preko 30.000 samo obaveštenje;
   - UVEK / PRAG: nivo po iznosu. Red u `Motor-Odobrenja` sa korakom, `NaRedu` (članovi DL opštih poslova se čitaju
     svaki put: *Office 365 Groups › List group members* **[proveriti]** za distribucionu listu, a ako ne radi, ručna
     lista u `Motor-Podesavanja`) i rokom (`rokRadnihSati`). Mejlovi za odobrenje → `Motor-IzlazniMejlovi`, a službi
     „čeka odobrenje".
3. `ODOBRENO` / `ODBIJENO` / `DOPUNA`:
   - **odluka je prva dozvoljena** ako `Odobrenje = Čeka odobrenje` i pošiljalac je u `NaRedu` (ili je rukovodilac /
     zamenik na nivou 1). Tada se upisuju polja odobrenja u zahtev, red u `Motor-Odobrenja` se zatvara i šalju se
     mejlovi;
   - nedozvoljena adresa → odgovor „na redu je …";
   - kasna odluka → odgovor „odlučio je … u …".
4. `DOPUNJENO` / nova `PROCENA` posle dopune: preko 200.000 prelazi na nivo 2.

Concurrency 1 i provera stanja pre upisa obezbeđuju pravilo „važi prva odluka" i kad dva odobravaoca odgovore u istom
minutu.

### 2.2 Grana tiketa (u F1)

1. Čišćenje teksta: *Html to text* (Content Conversion), pa `ocistiOdgovor` (odseca citat od prve linije „From: / Od: /
   Šalje: / -----Original", potpis posle `-- ` i uobičajene potpise). Pravila se prenose iz postojećeg koda.
2. Komentar: *Send an HTTP request to SharePoint*, POST na `.../items(<id>)/Comments` sa „Ime, vreme: tekst".
3. Slike iz mejla veće od 8 KB → *Add attachment* na zahtev.
4. Status: po naslovu (dugme) ili prvoj reči. Važi samo ako je pošiljalac iz našeg domena i Status je Otvoreno / U toku.
   Upisuju se `Status`, `DatumResenja` i `Odgovor`.
5. Pitanje (`?`) od nekoga ko nije domaćin:
   - domaćin u tenantu: komentar sa `mentions`, pa obaveštenje šalje SharePoint;
   - spoljni domaćin: `[Z-id] Pitanje službe` → `Motor-IzlazniMejlovi`.
6. Promena statusa → stavka u dnevnom pregledu (tebi, jednom posle rada F4 u 17 h, samo ako ima promena). Ovo zamenjuje
   pregled posle svakog rada.

**Tvoji odgovori iz Poslatih:** čitanje ličnih Poslatih se ukida. Pravilo je da u niti uvek bude `domacin@`, a to važi
automatski jer mejl službi stiže od `domacin@`, pa „Odgovori svima" ga uključuje. Ti odgovaraš iz zajedničkog sandučeta
(imaš Full Access) ili staviš `domacin@` u Cc.

---

## 3. Ulazna pošta bez tvog Inboxa

**Preporuka: zajedničko sanduče sa adresom `domacin@`.** M365 grupa ostaje vlasnik sajta i ništa se na sajtu ne menja.

- Power Automate nema pouzdan okidač za poštu **M365 grupe**. Outlook konektor podržava lično i **zajedničko** sanduče.
  Čitanje razgovora grupe preko Grapha traži aplikacione dozvole na celu grupu ili tenant (`Group.Read.All`) i Premium
  HTTP. Zato grupa ne.
- Postupak (radi IT):
  1. Primarna adresa grupe se promeni u npr. `domacin-grupa@`. Sajt i URL `/sites/domacin` se ne menjaju.
  2. Napravi se zajedničko sanduče **„Domacin radnje" `domacin@`**. Domaćini van tenanta nastavljaju da šalju na istu
     adresu, a zajedničko sanduče prima spoljnu poštu bez dodatnog podešavanja.
  3. Servisni nalog dobija **Full Access + Send As**, a ti Full Access (za pregled i ručne odgovore).
- **Prelaz za stare mejlove:** dugmad u već poslatim mejlovima vode na tvoju ličnu adresu. Na tvom sandučetu se pravi
  **serverska** Inbox pravila (bez računara): naslov sadrži `[Z-` → *Redirect* na `domacin@`. Ostaje uključeno dok ima
  otvorenih zahteva iz starog perioda, pa se gasi.
- Kopija svake poruke tebi više nije potrebna. Ako je želiš, to je opcija zajedničkog sandučeta, a motor je ne koristi.

---

## 4. Dugmad za status i odobrenje: `mailto` ostaje

| Opcija | Za | Protiv |
|---|---|---|
| **`mailto` na `domacin@`** (preporuka) | Radi u svakom klijentu i na telefonu. Bez registracije i licence. Naslov `[Z-id]` drži nit. Pravila „prva odluka", rokovi, dopuna i eskalacija su već definisana. Službe ništa novo ne uče. | Otvara nov mejl koji treba poslati (jedan klik više). Moguća greška u naslovu (motor odgovara „nije prepoznato"). |
| **Power Automate Approvals** | Dugmad direktno u Outlook-u, istorija u Approvals centru, standardno. | Mejl stiže od „Microsoft Power Automate", ne od `domacin@`, pa se kida nit = tiket. Jedno odobrenje = jedna odluka, pa za U toku → Rešeno treba više odobrenja. Nema prenosa na drugog odobravaoca posle roka, mora otkazivanje i novo odobrenje. DOPUNA i PROCENA se teško uklapaju. Treba Dataverse. |
| **Adaptive Cards (actionable messages)** | Najlepše iskustvo u Outlook-u. | Treba registracija provajdera u *Actionable Email Developer Dashboard* uz odobrenje admina tenanta i javni HTTPS endpoint (*When an HTTP request is received* = **Premium**). Radi samo u Outlook-u. Veći posao održavanja. |

**Zaključak:** u ovoj fazi `mailto` sa `domacin@`. Ako operativni direktor bude želeo, **Approvals se može uvesti samo za
nivo 2**: jedan odobravalac, rok 7 dana, bez prenosa, pa se dobro uklapa.

---

## 5. PDF izveštaja bez Edge-a

1. **Prvi izbor (standard):** HTML izveštaja (isti `izvestajHtml`, verzija za štampu, slike ugrađene kao `data:` URI,
   umanjene na 1000 px) → privremeni fajl u OneDrive servisnog naloga → **OneDrive for Business › Convert file** (format
   PDF) → SharePoint. **[proveriti]** Da li konverzija prikazuje ugrađene slike i kakav je prelom tabela. Ograničenje je
   veličina ulaznog fajla, a 15 slika × 300 KB je oko 6 MB u base64.
2. **Varijanta Word:** šablon `.docx` sa *Populate a Microsoft Word template* je **Premium** i ne radi dobro sa
   promenljivim brojem slika. Ne preporučuje se.
3. **Rezerva (Premium):** **Azure Function** (Node + Playwright/Chromium ili gotov kontejner), HTML ulaz → PDF izlaz,
   zaštićena ključem. Daje isti izgled kao današnji `--print-to-pdf`. Trošak na Consumption planu za oko 25 PDF-ova
   nedeljno je praktično nula, ali treba Azure pretplata i Premium HTTP za servisni nalog.

Predlog: kreni sa 1. Ako u probi slike ili prelom nisu prihvatljivi, pređi na 3.

---

## 6. Vlasništvo: da flow ne stane kad ti promeniš lozinku ili odeš na odmor

- **Servisni nalog** `svc-domacin@` **[P naziv]**:
  - **licenciran** (M365 licenca sa Power Automate za Office 365, npr. E3 ili Business Standard). Zajedničko sanduče
    nema licencu i ne može da se prijavi, pa **ne može** biti vlasnik konekcija;
  - vlasnik svih flow-ova i **svih konekcija**. Ti si **ko-vlasnik** flow-ova (uređuješ, gledaš istoriju, uključuješ i
    isključuješ);
  - član sajta `/sites/domacin` sa pravom Edit, a na listi `DomacinZahtevi` i bibliotekama dovoljno je Edit;
  - Full Access + Send As na `domacin@`.
- **Šta prekida konekcije i kako se sprečava:**
  - promena lozinke ili poništavanje sesija servisnog naloga → lozinka bez isteka, a promenu najavljuje IT i odmah
    obnavlja konekcije;
  - MFA i Conditional Access sa čestim ponovnim prijavama → IT pravi izuzetak ili posebnu politiku za servisni nalog
    (npr. prijava samo sa mreže firme), uz jaku lozinku u sefu IT-a;
  - konekcija nekorišćena 90 dana → ne preti, flow-ovi rade svaki sat.
- **Zamena:** drugi ko-vlasnik (npr. kolega iz operative) i kratko uputstvo u `uputstvo/` (kako se obnavlja konekcija,
  kako se prebacuje PROBA / ZVANIČNO, kako se pauzira).
- **Rezervna kopija:** izvoz rešenja (`.zip`) jednom mesečno u `Documents/Domacin radnje/motor/`.
- **Pauza posle probnog rada:** svi flow-ovi se isključe (*Turn off*), a liste i fajlovi ostaju. Ponovno uključivanje
  traje nekoliko minuta.

---

## 7. Licence i obim

### 7.1 Šta je standardno (pokriva M365 licenca servisnog naloga)

SharePoint (uključujući *Send an HTTP request to SharePoint*), Office 365 Outlook (zajedničko sanduče), Office 365
Users, Office 365 Groups, Excel Online (Business) *Run script*, OneDrive for Business *Convert file*, Content
Conversion, Approvals (ako se uvede), Schedule / Recurrence, child flow u rešenju.

### 7.2 Šta traži Premium

| Stavka | Kada je potrebna |
|---|---|
| HTTP / Azure Functions konektor | samo ako PDF ide preko Azure Function-a |
| *When an HTTP request is received* | samo za Adaptive Cards (ne preporučuje se) |
| Word Online *Populate template* | ne preporučuje se |
| Dataverse tabele (umesto SharePoint lista) | nije potrebno |

**Procena:** sa preporučenim rešenjem **dodatna licenca nije potrebna**, osim M365 licence za servisni nalog. Ako PDF
mora preko Azure-a: **Power Automate Premium za jednog korisnika (servisni nalog)** i Azure pretplata (trošak funkcije je
zanemarljiv).

### 7.3 Limiti za ovaj obim

- Obim: oko 47 poseta na 2 nedelje, što je oko 5 poseta dnevno, sa najviše 15 zahteva i 15 slika po poseti, plus
  nekoliko desetina odgovora i komandi nedeljno.
- Gruba procena: najviše **1.500–3.000 akcija dnevno** (najviše troše kreiranje zahteva i dodavanje priloga). M365
  licenca daje **6.000 Power Platform zahteva za 24 h po korisniku**, a Premium **40.000**. Ima prostora, ali je bolje
  koristiti *Select* / *Filter array* umesto *Apply to each* nad 90 stavki.
- Konektori: SharePoint i Outlook imaju ograničenja od nekoliko stotina poziva u minutu po konekciji. Ovaj obim je daleko
  ispod toga.
- Office Scripts: oko 1.600 poziva *Run script* dnevno po korisniku, 120 s po pozivu, oko 5 MB po pozivu **[proveriti
  aktuelne vrednosti]**. Dovoljno.
- Veličina mejla: zvanični izveštaj sa 15 slika je oko 6–8 MB. Proveriti limit Outlook konektora i Exchange-a u tenantu
  **[proveriti]**.

---

## 8. Plan prelaska bez prekida

Pravilo: u svakom trenutku **samo jedan motor upisuje i šalje za određeni korak**. Nov korak prvo radi u režimu **SENKA**:
čita stvarne ulaze, a umesto upisa i slanja beleži u `Motor-Log` šta bi uradio, dok mejlovi idu samo tebi (PROBA). Kad
se 3–5 radnih dana poklapa sa starim motorom, uključuje se, a odgovarajući korak u starom motoru se gasi.

| Faza | Šta se uključuje u novom | Provera | Šta se gasi u starom |
|---|---|---|---|
| **0. Priprema** (1–2 nedelje, zavisi od IT-a) | Servisni nalog, zajedničko sanduče (još bez adrese `domacin@`, privremeno `domacin-novi@`), liste `Motor-*` popunjene iz lokalnih CSV/JSON fajlova (jednokratni uvoz), F0 u PROBI, F6 | F0 šalje probni mejl kao `domacin-novi@` | ništa |
| **1. Podsetnici** | F4 c) i d) u SENKA/PROBA za 1–2 ponedeljka, pa zvanično | Isti broj mejlova i iste adrese kao stari ponedeljni podsetnik | korak „podsetnici". `POSLEDNJI_PONEDELJAK` se prenosi da ne ode dupli podsetnik |
| **2. Prebacivanje adrese** | IT premešta adresu `domacin@` sa grupe na zajedničko sanduče. Serverska pravila na tvom sandučetu (`[Z-` → redirect). Na zajedničkom sandučetu se pravi serversko pravilo „kopija svake poruke na moju adresu", kako bi stari motor i dalje
video svu poštu dok njegovi koraci ne budu ugašeni. F1 samo snima poruke i premešta ih u `Obrađeno`, bez obrade. | Spoljni test mejl stiže u zajedničko sanduče i kopija stiže tebi | ništa. Stari motor šalje kao zajedničko sanduče (IT daje tvom nalogu Send As do kraja prelaska), a čita kopije iz tvog Inboxa |
| **3. Tiket iz mejla** | F1 grana tiketa: SENKA 3 dana, pa zvanično | Komentari i statusi isti kao u starom (stari radi nad kopijom pošte. Ako to nije izvodljivo, SENKA se poredi ručno na 10–20 poruka) | korak „statusi" i čitanje Inboxa / Poslatih za `[Z-id]` |
| **4. Prijem poseta + izveštaj + zahtevi** | F2 + F5 + F1 grana „poseta mejlom". Novi mejlovi službi imaju dugmad na `domacin@`. | Prva poseta posle prelaska: proveriti broj zahteva, priloge, izveštaj i PDF | korak „registar (posete, izveštaj, zahtevi)" i lokalni OneDrive sync za motor |
| **5. Odobravanje** | F1 grana komandi + F4 a) i b). Otvorena odobrenja se prenose iz lokalnog JSON-a u `Motor-Odobrenja` u istom danu kad se stari korak gasi. | Jedno odobrenje nivoa 1 i jedno sa istekom roka, u PROBI | korak „odobravanje" |
| **6. Gašenje** | sve u ZVANIČNO | 2 nedelje bez intervencije | Task Scheduler zadatak „Collab - domacin radnje", Send As tvog naloga na `domacin@`, sertifikat aplikacije samo za analitiku (ostaje) |

Odobravanje je namerno poslednje jer ima najviše stanja i najveću posledicu greške. Podsetnici su prvi jer samo čitaju
listu i šalju.

---

## 9. Rizici i ograničenja

| Rizik | Mera |
|---|---|
| **Dupla obrada** (stari i novi motor istovremeno, ponovljeno pokretanje, ista poruka dva puta) | Prelazak korak po korak (8). Jedinstvene vrednosti u `Motor-Poruke` / `Motor-Posete`. Concurrency 1 na F1 i F2. Provera stanja pre upisa. |
| **Fajl posete stigne pre slika** | F2 čeka da stignu sve slike navedene u JSON-u, najviše 3 sata, pa šalje upozorenje. |
| **Upis motora menja `Modified` zahteva** | Stanje motora je u odvojenim listama (1.2). Prilozi iz mejla i statusi menjaju `Modified`, ali to su stvarne promene. |
| **Petlje okidača** (flow menja red, pa okine drugi flow) | Na listi `DomacinZahtevi` nema okidača *When an item is modified*. Procene upisane u listi se proveravaju po rasporedu (F4 a). |
| **Komentari i @oznaka** nemaju standardnu akciju | *Send an HTTP request to SharePoint* (standard). Proveriti format `mentions` i obaveštenje za nalog koji nije domaćinov. |
| **Prilozi stavki** | Standardna akcija *Add attachment*, jedan poziv po slici. Kod 15 slika × više zahteva troši akcije, što je uračunato u procenu u 7.3. |
| **Odgovor bez `[Z-id]`** (FW sa novim naslovom) | Ostaje poznata rupa. Ublažava se tako što F1 takve poruke iz našeg domena šalje u dnevni pregled kao „neprepoznato", pa ih ti ručno povežeš. |
| **Praznici** | Lista `Motor-Praznici`, koju `rokRadnihSati` uzima u obzir od prvog dana. Popuni se do kraja 2026. |
| **Vremenska zona i letnje vreme** | *Recurrence* sa zonom „Central Europe Standard Time", a datumi se računaju sa `convertTimeZone`. Proveriti prvi rad posle 25.10.2026 (prelazak na zimsko vreme). |
| **Pad konektora ili istek konekcije** | F6 i failure obaveštenja. *Retry policy* na akcijama slanja i upisa. Neposlati mejlovi ostaju u `Motor-IzlazniMejlovi` i šalju se ponovo. |
| **Gde ostaje log** | Istorija pokretanja flow-a čuva se 28 dana. Trajni log je u `Motor-Log` i `Motor-IzlazniMejlovi` (šta je, kome i kada poslato), sa verzijama liste. |
| **Distribuciona lista opštih poslova** | Ako *List group members* ne čita DL, članovi se drže u `Motor-Podesavanja` (ručno), ili IT pretvara DL u M365 / mail-enabled security grupu. |
| **Izvod iz ugovora o zakupu** | Lokalna baza više nije dostupna motoru. Potreban je jednokratni izvoz u `Motor-UgovoriFasada` i dogovor ko ga ažurira. |
| **Složenost u Power Automate-u** | Logika se drži u Office Scripts (jedan fajl, verzije u SharePoint-u), a flow-ovi ostaju tanki. Kod Office Scripts-a se čuva i u ovom repozitorijumu. |

---

## 10. Spisak za IT

| # | Zahtev | Obrazloženje |
|---|---|---|
| 1 | **Servisni nalog** `svc-domacin@…` sa **M365 licencom koja uključuje Power Automate** (npr. E3 / Business Standard) | Vlasnik flow-ova i konekcija, da motor ne zavisi od jedne osobe. Zajedničko sanduče ne može biti vlasnik konekcija. |
| 2 | Za servisni nalog: **lozinka bez isteka** i **Conditional Access / MFA politika** koja ne poništava sesije (npr. dozvoljena prijava samo sa mreže firme), uz čuvanje lozinke u sefu IT-a | Svako poništavanje sesije ili promena lozinke zaustavlja sve flow-ove dok se konekcije ne obnove. |
| 3 | **Zajedničko sanduče „Domacin radnje" sa adresom `domacin@`**. Adresa se premešta sa M365 grupe, a grupa dobija npr. `domacin-grupa@`. Prijem spoljne pošte uključen. | Power Automate čita i šalje iz zajedničkog sandučeta standardnim konektorom, a iz pošte grupe ne može. Spoljni domaćini ne menjaju adresu. Sajt `/sites/domacin` ostaje isti. |
| 4 | **Full Access + Send As** na `domacin@` za servisni nalog, i **Full Access** (+ privremeno Send As) za mene | Motor šalje kao „Domacin radnje". Ja pratim sanduče i ručno odgovaram, a stari motor radi do kraja prelaska. |
| 5 | Dozvola za **serverska Inbox pravila sa preusmeravanjem** (*Redirect*) **unutar tenanta** na mom sandučetu | Odgovori na stare mejlove sa dugmadima na moju adresu stižu u motor bez mog računara. Ako je auto-forward blokiran politikom, potreban je izuzetak za internu adresu. |
| 6 | **Power Platform okruženje sa Dataverse-om** (ili dozvola za rešenja u postojećem) i **DLP politika** koja dozvoljava zajedno: SharePoint, Office 365 Outlook, Office 365 Users, Office 365 Groups, Excel Online (Business), OneDrive for Business, Content Conversion | Rešenje (solution) omogućava child flow-ove, izvoz i prenos. DLP ne sme da razdvoji ove konektore u različite grupe. |
| 7 | **Office Scripts uključeni** za tenant (Microsoft 365 admin › Settings › Office Scripts) i dozvola da se pokreću iz Power Automate-a | Logika rokova, čišćenja odgovora i izveštaja ide u Office Scripts. |
| 8 | Servisni nalog kao **član sajta `/sites/domacin`** (Edit) i pravljenje novih lista `Motor-*` na sajtu | Upis zahteva, priloga i komentara, i stanje motora umesto lokalnih fajlova. |
| 9 | Potvrda da **Office 365 Groups › List group members** čita **distribucionu listu opštih poslova**, ili njeno pretvaranje u M365 / mail-enabled security grupu | Odobravaoci nivoa 1 se čitaju pri svakom radu. |
| 10 | *Uslovno, samo ako konverzija PDF-a preko OneDrive-a nije dobra:* **Power Automate Premium** za servisni nalog i **Azure pretplata / resource group** za jednu Azure Function (Consumption) | PDF izveštaja sa ugrađenim slikama bez Edge-a na mom računaru. Trošak funkcije je zanemarljiv. |
| 11 | *Nije potrebno u ovoj fazi:* registracija **actionable messages** provajdera | Predlog zadržava `mailto` dugmad (odeljak 4). Biće potrebna samo ako se kasnije uvedu Adaptive Cards. |
| 12 | *Informativno:* Entra aplikacija sa `Sites.Selected` ostaje samo za moju lokalnu analitiku | Motor u oblaku je ne koristi, pa nema novih dozvola za aplikaciju. |

---

## Pretpostavke

- Tenant ima M365 E3 ili sličnu licencu, sa standardnim Power Automate-om za Office 365.
- Svi zaposleni u službama imaju sandučad u našem domenu. Domaćini van tenanta komuniciraju samo mejlom.
- Imena kolona i postojeće liste se ne menjaju, nove su samo liste `Motor-*`.
- Izvod iz ugovora o zakupu se može jednom izvesti u SharePoint listu.
- Format naziva perioda u `izvestaji/<period>` ostaje kao danas. Period počinje 22.6.2026 + 14·n dana, a vikend posle
  petka spada u isti period.
