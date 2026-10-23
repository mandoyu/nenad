# Izveštaj „promet u odnosu na budžet": iz 22 h prebaciti na jutro, sa podacima iz baze

Stanje: provera izvodljivosti, oktobar 2026. Pretpostavke su označene sa **[P]**. Stavke koje se moraju potvrditi upitom
nad bazom ili kod vlasnika izveštaja označene su sa **[proveriti]**. Ova provera je urađena bez pristupa bazi i botu, pa
brojke o veličini propusta još ne postoje. Odeljak 5 daje upite kojima se izmere za jedan dan rada.

---

## 0. Ukratko

- **Izvodljivo je, i bolje je od sadašnjeg.** Collab REP (MariaDB, ODBC DSN `Collab REP`) je replika bez kašnjenja.
  Jutarnji upit u 07:00 vidi ceo prethodni dan, uključujući web posle 22 h, povrate i stornacije. Bot koji čita ekran
  Collaba više nije potreban.
- **Sadašnji izveštaj gotovo sigurno potcenjuje dan.** Web (WP01, P201, P202) prodaje do ponoći, a izveštaj se zatvara
  u 22 h. Da li se taj deo pojavi sutradan u kumulativu meseca zavisi od toga šta tačno bot čita **[proveriti]**. Ali
  poslednjeg dana u mesecu prodaja 22–24 h sigurno ne ulazi u mesec u kom je nastala.
- **Glavni posao nije zakazivanje nego pravilo avansa.** Prodaja se priznaje na dan izdavanja avansnog računa, isto kao
  danas. To u bazi znači tri vrste dokumenata (avans, storno avansa, konačni račun) i jedno pravilo protiv duplog
  brojanja (odeljak 3). `TM_SALES`, osnova za analitiku, **ne radi po tom pravilu** **[P]**, pa ga ne treba koristiti
  za ovaj izveštaj bez dorade.
- **Dve zamke u opsegu:** B2B posle 01.04.2025 nije u Collabu (prešao u SAP), a MTS postoji tek od 15.05.2026. Opseg
  kanala u bazi mora tačno da odgovara opsegu budžeta, inače izveštaj ili fali ili duplira (odeljak 4).
- **Preporuka:** 2–3 nedelje paralelnog rada (bot u 22 h + novi upit ujutru), zatim gašenje bota. Novi izveštaj radi
  kao PowerShell skripta u Task Scheduler-u, isto kao `snimi_zalihe.ps1` (06:00), i šalje mejl u 07:30.

---

## 1. Šta nije u redu sa sadašnjim izveštajem

| # | Problem | Posledica | Sigurno / verovatno |
|---|---|---|---|
| 1 | **Zatvaranje u 22 h**, a web radi do 24 h. Radnje u tržnim centrima koje rade do 22 h izdaju poslednje račune i nekoliko minuta posle 22 h. | Dnevni promet je potcenjen za deo web prodaje u 22–24 h i za poslednje račune radnji. | sigurno za dnevni broj |
| 2 | **Kraj meseca:** prodaja 1.–31. u 22–24 h | Ako bot čita kumulativ meseca, ovaj deo se vidi sutradan, ali 31. u 22–24 h ulazi u **sledeći mesec ili nigde**. Mesečna realizacija budžeta je pogrešna baš u izveštaju koji se najviše gleda. | sigurno za poslednji dan |
| 3 | **Stornacije i povrati** | Ne zna se da li ekran Collaba koji bot čita oduzima storno avansa i povrate. Storno urađen sutradan za avans od danas sigurno ne menja već poslat broj. | **[proveriti]** |
| 4 | **Bot čita ekran (RPA)** | Lomi se na svaku promenu ekrana, lozinke ili rezolucije. Ne ostavlja trag šta je pročitao. Brojka se ne može ponoviti ni proveriti. | sigurno |
| 5 | **Izveštaj u 22 h** stiže kad ga niko ne čita do jutra | Prednost „istog dana" je mala, a cena je nepotpun broj. | procena |

Zaključak: jutarnji izveštaj za zatvoren dan rešava 1, 2 i 5. Upit nad bazom rešava 3 i 4.

---

## 2. Predlog

```
06:00  snimi_zalihe.ps1 (postoji)
07:00  promet_vs_budzet.ps1
         ├─ ODBC Collab REP: prodaja po pravilu avansa za dan D-1 i mesec do D-1
         ├─ budžet (Excel / CSV, po danu × kanal × baner)          [proveriti izvor]
         ├─ kontrole (odeljak 6) — ako padne, mejl ide samo meni
         └─ upis rezultata u CSV (trag šta je poslato)
07:30  mejl primaocima današnjeg izveštaja
```

- **Termin 07:00.** Replika kasni 0 minuta, pa ništa ne čeka. Dovoljno je posle ponoći i posle noćnih obrada u Collabu,
  ako ih ima **[proveriti]**. 07:30 ostavlja pola sata da se kontrole pogledaju pre slanja.
- **Sadržaj (predlog):** juče (realizacija / budžet / %), mesec do juče (isto), ostatak do kraja meseca (potreban dnevni
  prosek za 100 %). Preseci: kanal (MPO / Web / CC / MTS, B2B samo ako je u budžetu) i baner (Tehnomanija, WinWin,
  Macola, Dijaspora). Isti preseci i poredak kao u današnjem izveštaju, da se primaoci ne navikavaju na nov izgled
  **[proveriti današnji izgled]**.
- **Prvog u mesecu** izveštaj donosi konačan prethodni mesec. To je prvi put da mesečna realizacija ima ceo poslednji
  dan.
- **Linija opsega** na vrhu svakog izveštaja, npr.:
  `Opseg: svi baneri | MPO + Web + CC + MTS (bez B2B) | 01–22.10.2026 | promet sa PDV | prodaja = izdat avans`
- **Gde radi:** na mom računaru, kao i ostale skripte nad bazom. Prebacivanje u oblak (Power Automate, vidi
  `domacin-radnje-motor-u-oblaku.md`) **nije preporučljivo za ovaj izveštaj**: Collab REP je u lokalnoj mreži, a
  pristup iz Power Automate-a traži *on-premises data gateway* i Premium konektor. Ako računar mora da otpadne, prava
  zamena je serverski Task Scheduler na mašini u mreži firme koja već ima ODBC pristup **[proveriti sa IT-om]**.
- **Večernji „flash" u 22 h** se može zadržati kao „preliminarno, bez web prodaje 22–24 h", ali iz baze, ne od bota.
  Predlog je da se ukine ako ga niko ne traži posle prve nedelje.

---

## 3. Pravilo „prodaja = izdat avans" u bazi

Današnji izveštaj smatra prodajom dana **izdat avans**. Novi mora da daje isti broj, inače se poređenje sa budžetom i
sa prošlim mesecima kida. Iz fiskalizacije znamo oblik **[P]**: web i CC porudžbina dobija **avansni račun** kad kupac
plati (ili kad se porudžbina potvrdi), a **konačni račun** kad se roba isporuči ili preuzme. Konačni račun umanjuje se za
već plaćen avans.

### 3.1 Šta se broji i kog dana

| Dokument | Dan kom pripada | Iznos |
|---|---|---|
| **Avansni račun** | dan izdavanja | + iznos avansa |
| **Storno / refundacija avansa** (otkazana porudžbina) | **dan storna** | − iznos avansa |
| **Konačni račun bez avansa** (obična kupovina u radnji) | dan izdavanja | + ceo iznos |
| **Konačni račun sa avansom** | dan izdavanja | + samo **doplata** (iznos − avans). Ako je avans pokrio sve, 0. |
| **Povrat robe / knjižno odobrenje** | dan povrata | − iznos |
| **MTS nalog** (`MTSNZO-`) | dan odobrenja naloga | + iznos (MTS nema avans) **[proveriti]** |

Dve odluke koje treba doneti sa vlasnikom izveštaja **pre** pisanja upita:

1. **Storno avansa ide na dan storna ili na dan avansa?** Predlog: na dan storna. Tada se već poslat dan nikad ne menja
   i zbir meseca je tačan. Ako se storno vraća na dan avansa, svaki izveštaj menja prošle dane, a mesec se „pomera"
   posle zatvaranja.
2. **Delimičan avans.** Ako kupac plati 10 % unapred, da li se prodajom smatra 10 % na dan avansa i 90 % na dan
   isporuke (predlog, sledi iz tabele) ili ceo iznos porudžbine na dan avansa? Odgovor zavisi od toga kako radi
   današnji izveštaj **[proveriti kod vlasnika i na jednom stvarnom primeru]**.

### 3.2 Šta treba naći u bazi **[proveriti]**

- **Doktip avansnog računa** i **doktip storna avansa** u `dokumenti`. U `TM_SALES` su poznati doktipovi 10022
  (maloprodaja), 10019 (`FPD-`, B2B), 10115 (`KFA-`), 10257 / 10259 (MTS). Avans među njima verovatno **nije**, jer je
  `TM_SALES` pravljen za analitiku po isporuci **[P]**.
- **Veza konačnog računa sa avansom** (kolona u `dokumenti` ili posebna tabela). Bez nje se doplata ne može izračunati i
  web prodaja se broji dvaput: jednom kao avans, jednom kao konačni račun.
- **Da li `dokdatum` nosi vreme.** Za izveštaj nije bitno (gleda se ceo dan), ali jeste za merenje propusta 22–24 h
  (odeljak 5).
- **PDV na avansu.** Avansni račun ima PDV, pa je promet sa PDV-om direktan. Ako je budžet u prihodu (bez PDV), uzima
  se osnovica sa avansa **[proveriti u kojoj meri je budžet]**.

### 3.3 Posledica za analitiku

`TM_SALES` i izveštaji iz njega priznaju web prodaju na dan konačnog računa **[P]**. Novi izveštaj priznaje je na dan
avansa. Dnevni brojevi se zato neće poklapati, a mesečni samo približno (porudžbine plaćene krajem meseca, isporučene
početkom sledećeg). To nije greška, ali mora pisati u napomeni izveštaja, inače će neko uporediti dve brojke i tražiti
razliku.

---

## 4. Opseg kanala: šta u bazi ima, a šta nema

| Kanal | U Collab REP | Napomena |
|---|---|---|
| MPO | da | `tip_lokacije = 'Radnja'`. P097 (outlet), P537 i P095 se **ne izuzimaju** iz prometa, jer su u budžetu **[proveriti]**. Izuzimaju se samo iz analize RUC-a. |
| Web | da | WP01, P201, P202. Web sa preuzimanjem u radnji je Web, ne MPO. |
| CC | da | WP02, P182, P183 |
| MTS | da, od 15.05.2026 | prefiks `MTSNZO-`. Atribucija po „ko je prodao", ne po DC14. |
| **B2B** | **ne posle 01.04.2025** | B2B je u SAP-u. Ako je B2B u budžetu, mora se dodati iz SAP-a ili izveštaj mora reći „bez B2B". Ako se dodaje, **bez Telekom B2B (partner 20001163)**, inače se MTS broji dvaput (~1,21 mlrd RSD rizika). |
| Prijemna `…V` i DC šifre | da, ali nisu prodaja | isključiti (`tip_lokacije = 'Prijem'`, DC) |

Baner se uzima iz `d_prodavnica.baner`, ne iz prefiksa šifre. Kanalske šifre (WP01…P202) nisu u MPO Excelu, pa
spajanje po baneru mora ići preko dimenzije, inače ~10 % prometa tiho ispada.

---

## 5. Kako izmeriti koliko sadašnji izveštaj greši (pre odluke)

Tri upita, jedan dan posla, daju brojke umesto procene. Imena kolona su okvirna i biraju se kad se potvrdi odeljak 3.2.

**5.1 Prodaja posle 22 h po kanalu, za poslednja 3 meseca**

```sql
-- udeo prometa izdatog posle 22:00, po kanalu i danu u nedelji
SELECT m.sifra,
       DAYOFWEEK(d.dokdatum)                      AS dan,
       SUM(<promet>)                              AS promet_ukupno,
       SUM(CASE WHEN TIME(d.dokdatum) >= '22:00:00' THEN <promet> END) AS promet_posle_22
FROM dokumenti d
JOIN magacini m ON m.idmagacin = <magacin dokumenta>
WHERE d.doktip IN (<avans>, <storno avansa>, 10022, 10257, 10259)
  AND d.dokdatum >= '2026-07-01' AND d.dokdatum < '2026-10-01'
GROUP BY m.sifra, DAYOFWEEK(d.dokdatum);
```

Očekujem da je web jedini kanal sa primetnim udelom, sa najvećim udelom nedeljom i ponedeljkom uveče **[P]**.

**5.2 Poslednji dani u mesecu:** isti upit samo za 31.07., 31.08. i 30.09.2026. To je iznos koji je mesečna realizacija
budžeta izgubila u ta tri meseca.

**5.3 Usklađivanje sa botom:** za 10 poslednjih dana izračunati promet po pravilu iz 3.1 **sa granicom 22:00** i
uporediti sa brojem koji je bot poslao. Ako se poklapa, pravilo avansa je dobro prepoznato i preostaje samo pomeranje
granice. Ako se ne poklapa, razlika pokazuje šta bot zapravo čita (npr. bez storna, ili bez MTS-a).

---

## 6. Kontrole pre slanja (svakog jutra)

Ako bilo koja padne, izveštaj ide samo meni sa opisom, a primaocima ništa (bolje bez izveštaja nego pogrešan).

| Kontrola | Uslov |
|---|---|
| Svežina | postoji bar jedan dokument posle 23:00 juče (`AND dokdatum <= NOW()`) |
| Pokrivenost dimenzije | zbir po baneru = ukupan zbir (nova šifra mora u `d_prodavnica` pre izveštaja) |
| Avans bez veze | konačni računi sa avansom kojima veza ka avansu nije nađena = 0 |
| Skok | dan odstupa više od ±50 % od istog dana prošle nedelje → upozorenje, ne blokada |
| Budžet | postoji red budžeta za juče za svaki kanal × baner |

---

## 7. Plan

| Korak | Šta | Ko | Trajanje |
|---|---|---|---|
| 1 | Doktip avansa i storna, veza konačni račun → avans, `dokdatum` sa vremenom (3.2) | ja, upitom nad bazom | 1 dan |
| 2 | Odluke iz 3.1 (storno, delimičan avans), izvor i mera budžeta, da li je B2B u budžetu | vlasnik izveštaja | 1 sastanak |
| 3 | Merenje propusta (odeljak 5) | ja | 1 dan |
| 4 | Skripta `promet_vs_budzet.ps1` + kontrole, mejl samo meni | ja | 2–3 dana |
| 5 | Paralelni rad: bot u 22 h ostaje, nova skripta ujutru meni. Svaki dan poređenje sa granicom 22:00 (mora da se poklopi) i bez granice (razlika = propust). | ja | 2–3 nedelje, uključujući jedan kraj meseca |
| 6 | Najava primaocima: novo vreme, novi broj, napomena o razlici prema `TM_SALES` | vlasnik izveštaja | — |
| 7 | Prebacivanje primalaca na jutarnji izveštaj, gašenje bota | ja + ko održava bota | 1 dan |

Paralelni rad treba da obuhvati bar jedan kraj meseca (31.10. ili 30.11.2026), jer se tu vidi najveća razlika.

---

## 8. Otvorena pitanja za vlasnika izveštaja

1. Šta tačno bot čita u Collabu (koji ekran / izveštaj, dan ili kumulativ meseca)? Da li su storna u tom broju?
2. Storno avansa: dan storna (predlog) ili dan avansa?
3. Delimičan avans: deo na dan avansa, ostatak na isporuci (predlog), ili cela porudžbina na dan avansa?
4. Budžet: u kom je obliku i gde (fajl, granularnost dan / mesec, kanal × baner × radnja), u prometu ili prihodu?
5. Da li su B2B i MTS u budžetu? Ako je B2B, da li ga dodajemo iz SAP-a ili izveštaj ide „bez B2B"?
6. Ko prima izveštaj i da li iko stvarno koristi verziju u 22 h?
7. Prelazak na zimsko vreme 25.10.2026: Task Scheduler radi po lokalnom vremenu, pa nema uticaja na jutarnji termin.
   Bot u 22 h, ako radi po UTC-u, mogao bi se pomeriti za sat **[proveriti]**.
