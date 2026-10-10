PERMIT SERVER - all permit tools as ONE web application hosted on YOUR PC
=========================================================================
Your PC is the server. The database and every file stay on your PC. Colleagues open it in Chrome / Edge -
nothing to install on their PCs, no internet, no GitHub, no cloud.

  Danfoss  |  Festo  |  Agilent  |  MSD (export)  |  Any supplier   ->   dec-upload Excel + CHECK_REPORT + FrtImport XML + RAW_DATA
  (the same reading rules as the desktop tools, the same UPDATE HS CODE list and output format)


TRAINED LAYOUTS: every document trained so far (Sandvik, VF, Shell, Panasonic, Halliburton, Nurmi, DHL, DSV ...) is
listed in TRAINED_DOCUMENTS.txt (also the prompt to train a new one in another chat).

1. MAKE THE SERVER (one time, on the PC that will be the server)
   - Install Python 3.10+ from https://www.python.org/downloads/ (tick "Add python.exe to PATH").
   - Unzip this folder (e.g. to D:\PERMIT_SERVER - a drive that is backed up).
   - Double-click BUILD_SERVER_EXE.bat  ->  PERMIT_SERVER.exe appears (OCR + HS list + tools inside).
   - Right-click OPEN_FIREWALL_PORT.bat > Run as administrator (one time) - lets office PCs connect.
   (Without building: START_SERVER_WITHOUT_EXE.bat runs it with Python.)

2. START IT
   - Double-click PERMIT_SERVER.exe. A black window shows the addresses - KEEP IT OPEN (closing it stops the server).
     This PC:     http://localhost:8080
     Other PCs:   http://<this PC's IP>:8080     (shown in the window and on the Settings page)
   - First time: the browser asks you to create the admin account.
   - START_WITH_WINDOWS.bat: the server starts by itself when you log in to Windows.
   - The server PC must stay on (and not sleep) while people use it.
   - Different port: edit PERMIT_SERVER_DATA\server_config.json ("port": 8080) and run OPEN_FIREWALL_PORT.bat 8090.

3. USE IT (browser)
   Jobs        every job of every user - search by HBL, invoice, job no, user; status: Reading / To check / Done.
   New job     pick the supplier (or Auto), drop the PDFs (invoice, packing lists, B/L, customs instructions, air waybill,
               scans), optional job no / permit type / UEN / pre-alert text -> "Upload and read".
   One job     Items         the item page: click a cell to change HS / HS UOM / COO / CASC / description / net
                             weight (Enter = save + next row). Tick rows -> set one value on all of them, clear
                             your changes, keep as separate rows / combine, copy to Excel. Problems / controlled /
                             edited filters, search, sort.
               Item details  one item on a page: invoice line, every permit value with where it came from,
                             HS description, packing-list rows, notes. Save / Save and next / Next problem.
               Shipments     one Excel per HBL: use or skip, file name, HBL, importer, UEN, gross weight,
                             packages, other charges (invoice total only / add to last item).
               Header and cargo   permit type, job no, declaration type, transport, every header box (blank = from
                             the documents), remarks, pre-alert text, containers (TNE, rounded up, "Weight from
                             gross weight").
               Notes         all MISMATCH / MISSING / CHECK notes.
               Files         your PDFs (add more - the job is read again, edits kept) and the made files.
               Log           what the server did.
               Buttons: Make Excel + XML, Download files (zip), Undo, Read again, Copy job, Mark as done, Archive,
               Delete.
   HS codes    search the UPDATE HS CODE list (code or words): UOM, import / export control, CA, IDAMISC,
               product codes.  "Supplier HS to Singapore HS" mapping used by every job.
   Settings    server address, XML IDs (ITG_INDIA / YCH / IPTDEC), insurance, rules, forwarder names, location codes
               (PPZ, ME, KZ, JZ, TZ, CZ, O), product masters (download / upload), new UPDATE_HS_CODE.xlsx (upload -
               every tool uses it at once), backup.
   Users       (admin) add users, reset passwords, admin / user, switch off.
   Activity    (admin) who did what and when.

   Same rules as before: CASC = IDAMISC when the HS has it, otherwise MISC when controlled; discharge port not
   filled for imports; invoice sheet = total value of the shipment; container weight in TNE rounded up.

4. WHERE THE DATA IS (folder PERMIT_SERVER_DATA next to the exe)
   permit.db          the database (SQLite): users, jobs, file list, logs, settings, HS map, activity
   jobs\<no>\in       uploaded PDFs          jobs\<no>\out   Excel / XML / CHECK_REPORT / RAW_DATA
   jobs\<no>\state.pkl   everything read + all edits of that job
   masters\           product masters        hs_data.json    the HS code list in use
   Backup: Settings > Backup (database + masters, or everything incl. job files) - keep it on another drive.
   Moving to a new PC: copy the whole PERMIT_SERVER folder incl. PERMIT_SERVER_DATA.

5. ON THE INTERNET (open it from anywhere: home, mobile, other offices)
   A. YOUR PC STAYS THE SERVER + DATABASE  (recommended - free, no router / firewall change)
      - Double-click GO_ONLINE.bat   (or Settings > Internet address > "Go online").
        The first time it downloads cloudflared.exe (Cloudflare's free secure tunnel, about 60 MB).
      - The window and Settings page show the address, e.g. https://brave-cargo-xxxx.trycloudflare.com
        -> send it to your team. It is https (encrypted). Tick "go online every time the server starts".
      - This free address CHANGES every time the link starts.
        FREE FIXED LINK WITH GITHUB PAGES (recommended): your team always opens  https://<you>.github.io/permit-link/
        and it forwards to the current address. Only this forwarding page is on GitHub - jobs, PDFs and the database
        stay on your PC.
          1. github.com > New repository  permit-link  (Public, tick "Add a README").
          2. That repository > Settings > Pages > Source "Deploy from a branch", branch main, folder / (root) > Save.
          3. Your profile > Settings > Developer settings > Personal access tokens > Fine-grained tokens > Generate:
             Repository access "Only select repositories" = permit-link, Permissions > Contents = Read and write.
          4. Permit Server > Settings > Internet address > "Fixed link with GitHub Pages": type  <you>/permit-link
             and the token > "Save and update now".
          Every time the server goes online it updates the page (about 1 minute); when you switch it off the page says
          "offline".
        Or a fixed address on your own domain (e.g. permits.infinity.com.sg):
          1. Free Cloudflare account + your company domain on Cloudflare (dash.cloudflare.com).
          2. Zero Trust > Networks > Tunnels > Create tunnel (Cloudflared) > copy the TOKEN.
             Public hostname: permits.yourdomain.com  ->  Service  http://localhost:8080
          3. On the server PC: Settings > Internet address > "Your company domain": paste the token + the address > Save
             (or PERMIT_SERVER_DATA\server_config.json: "tunnel_token" / "public_url").
          4. Go online (or GO_ONLINE.bat). Tick "go online every time the server starts".
      - Works only while this PC is on, awake and connected. Switch it off: Settings > "Switch off" (on this PC).
   B. A CLOUD SERVER (runs 24 hours without your PC - needs a paid cloud account: any Linux VPS, Render, Railway,
      Fly.io ...). The Dockerfile + docker-compose.yml are included:
          docker compose up -d        (set PERMIT_ADMIN_PASSWORD in docker-compose.yml first)
      The database and files are then on that cloud server (volume "permitdata"); put HTTPS in front of it.

   SAFETY ON THE INTERNET
   - Anyone with the link sees the log-in page: give every person their own user + strong password (8+ characters),
     switch off people who leave (Users page). 8 wrong passwords = that log-in is blocked for 15 minutes.
   - The first admin can only be created on the server PC itself (http://localhost), never over the internet.
   - Everything travels encrypted (https). The Activity page shows every log-in, failed log-in and change.
   - Keep the Backup (Settings) on another drive.

7. RUN IT WITHOUT A WINDOW (in the background)
   - START_SERVER_BACKGROUND.bat (or PERMIT_SERVER.exe --background): starts the server with NO window and opens the
     browser. You can close the black window - the server keeps running. GO_ONLINE.bat does the same + internet link.
   - GO_ONLINE.bat while the server already runs: it switches the internet link on in the running server and shows the
     https address (no restart needed).
   - STOP_SERVER.bat stops it (or Settings > Server address > "Stop the server", admin on this PC only).
   - START_WITH_WINDOWS.bat: starts it in the background every time you log in to Windows.
   - Double-clicking PERMIT_SERVER.exe while it already runs just opens the browser (never two servers).
   - Log of the background server: PERMIT_SERVER_DATA\server.log

8. EVERY PERMIT HEADER FIELD
   - Header and cargo tab > "More permit fields": Previous Permit No., BG Indicator, Reference Documents, Shipment ID,
     Customer Remarks, Consignee, Supplier/Importer Relationship ... and "Add another field" = ANY column of the Header /
     Invoice sheet of the dec-upload template. Written into the Excel and the FrtImport XML. Tick "every job" (admin) to
     use a value on every new job.
   - Settings > Declarant: TradeNet Mailbox ID, declarant name, telephone, CR UEI - shown on every job for reference.
     The dec-upload Excel and the FrtImport XML have no fields for these: the permit system takes them from the declarant
     log-in (as on its Header page).

9. OTHER CHARGES ON THE INVOICE (e.g. Festo packing charges) - Shipments tab, per HBL
   - Add to last item           : last item of each invoice + charge; items = invoice total.
   - Invoice total only         : Invoice sheet = overall value, items = goods only (Validate shows it as a point to review).
   - Other Taxable Charge line  : Invoice sheet = goods value + "Other Taxable Charge" amount; XML gets an OTC line;
                                  items + charge = invoice total.

10. KNOWN LAYOUTS (Any supplier / Auto)
DALIAN NURMI HYDRAULICS (scanned pack to Sandvik Singapore) - read automatically
  One PDF = summary page ("Contact") + one INVOICE per order + Delivery Note.
  - The summary page repeats the invoices: it is NOT counted as items, only checked (summary total = sum of invoices).
  - Per invoice: invoice no. / date, PO, item code (OCR "O" -> "0" fixed), description, HS (10-digit -> 8), qty, price,
    amount, Package Fee (other charge - choose: add to last item / invoice total only / Other Taxable Charge).
  - Delivery note: net + gross weight per article -> each item; packages (39 wooden cases), total net / gross.
  - COO is not printed: CN (manufacturer in Dalian) is used and listed as CHECK.
SAFETY FOR EVERY SUPPLIER: the same invoice or packing list added twice (copy, full pack + split PDF) is used once.

6. GOOD TO KNOW
   - Jobs run one at a time (big scanned packs take a few minutes); others wait in the queue ("Waiting").
   - Only people on the same office network can open it. Do not open the port on the internet router.
   - Passwords are stored hashed. Log-ins last 12 hours.

11. MSD (EXPORT) - the MSD PERMIT TOOL inside the server
   New job > "MSD (export)" (Auto finds it from "Merck Sharp" / "MSD International" on the documents).
   - Documents: MSD invoices / tax invoices + packing lists + bill of lading - matched by Delivery # / order no / the
     B/L delivery list; all deliveries of one B/L = one Excel (Out HAWB = B/L No.).
   - DHL instruction e-mail: paste it in "Cargo details (pre-alert / e-mail)" or upload it with the PDFs (.txt / .eml, or
     a PDF of the e-mail). Used as in the desktop tool: CARRIER, CARRIER BOOKING REF (Outward MAWB), PLACE OF RELEASE
     ("O" + address), PLACE OF RECEIPT (PPZ), TYPE OF PERMIT (OD -> DRT), COO, HS CODE (check), EXPORTER + UEN,
     SHIPMENT TYPE FCL, JOB NO, PERMIT REMARKS.
   - Same Excel rules as the desktop tool (export fields only, description ITEM , DESC , QTY , NW , GW, pallets PAT,
     container weight TNE, Term Type not filled). Product master: Settings > Product masters > MSD (export).
   - Plus the server extras: Validate page, item page edits + undo, More permit fields, FrtImport XML, RAW_DATA.
   - CASC: IDAMISC when the HS has it, otherwise the controlled CASC code (Settings) for export-controlled HS codes.

12. NEW SUPPLIERS WITHOUT REBUILDING - plugins  (full steps + the prompts: HOW_TO_ADD_A_SUPPLIER.txt)
   Settings > Suppliers and readers: tools, built-in layouts, plugins (green = loaded, red = error + reason).
   - A plugin = one .py file (made in a new chat from PLUGIN_TEMPLATE.py + sample PDFs) in PERMIT_SERVER_DATA\plugins.
   - "Add plugin (.py)" (admin, on the server PC only) or copy the file there + "Reload". No exe rebuild, no restart.
   - "Test a PDF": what a reader makes of one PDF (reader used, invoice no, total, lines, HS, COO, weights, notes).
   - Plugins shipped in the server folder (PERMIT_SERVER\plugins, e.g. sandvik_tax_invoice.py) are copied into
     PERMIT_SERVER_DATA\plugins at start when not there yet (a plugin you changed there is never overwritten).
   - Plugins are tried before the built-in layouts, in "Any supplier" and "Auto" jobs.

13. THE JOB SCREENS (as the permit system)
   Jobs list: JOB NO. | DECL. DATE | DECL. TYPE | CREATE | ETA / ETD | PERMIT NO. | IMPORTER / EXPORTER | HAWB | MAWB/OBL |
     POL / POD | XREF (invoices) | ST | INT REMARKS | SHIPMENT ID. Click a column title to sort, tick jobs for Mark done / Archive.
     ST: DRF draft (to check), RDY Excel made, DON done, PMT permit no. entered, RD reading, QUE waiting, GEN making files,
     ERR error, ARC archived.
   One job: Job Information on the left (permit number after approval + internal remarks can be typed there) and one tab per
     sheet of the dec-upload Excel: Header | Invoice | Item | Casc | Container | Items to be consolidated, then Summary
     (validation, download and documents, notes, log) and Requirements (job / header / cargo rules, shipments, items check,
     item details, and the link to the rules for every job in Settings).
   - The sheet tabs show exactly what goes into the Excel (yellow = filled). Type in any field / cell: the value is written into
     that Excel and kept when the job is read again (orange = typed by you, the arrow puts it back to the documents / rules).
   - Item tab: COO, HS Code, HS UOM and CASC follow the item rules (a new HS code also changes HS UOM, HS qty and CASC);
     tick "Save ... to the product master" to remember them for the next shipment.
   - Several HBL = several Excel files: choose the permit at the top of the tab.

NEW KNOWN LAYOUTS (Any supplier / Auto) - October 2026
  SANDVIK SMC LOGISTICS (export from the Singapore hub - EXPORT permits only)
    INVOICE (Export HS Code + Country of Origin per line, net / gross weight per line) + Packing List (dispatch note).
    Item Description = "TOOLS OR PARTS INV NO : ( 92891039 )" on every item (fixed text + the invoice number).
      Typing a description on the Item tab replaces it for that item; untick the rule on Customer rules to use the
      invoice descriptions.
    Lines without Export HS Code: type the HS on the Item tab and tick "Save to product master" - next time it fills itself.
    Exporter "SANDVIK SMC LOGISTICS LTD": type the UEN once (remembered). HAWB comes later (booking) - type it.
  DHL CUSTOMS INSTRUCTIONS (BROKERAGE BSIN...): importer + UEN, packages + unit, gross weight, carrier agent, ocean B/L
    (MAWB/OBL), house B/L, vessel / voyage, port of loading, ETA / ATA (arrival date), place of release / receipt,
    type of permit (ME / IG / ...), container (no., 20GP, seal, FCL/LCL), permit remarks, freight (FRT + BAF -> Freight
    Charge when the invoice is FOB / FCA / EXW), DHL as freight forwarder + UEN.
  DSV ARRIVAL NOTICE + DSV SEA WAYBILL / BILL OF LADING: house B/L wins over the DSV "S" reference (SRAF0583645 = B/L
    RAF0583645), ocean B/L, vessel / voyage, ports, containers (CY/CY = FCL, CFS = LCL).
  SHANGHAI EVEREST FINE CHEMICALS invoice + packing list (Halliburton): lines per PO, qty in pounds, one HS (6 digits ->
    8 digits) + COO for the document, net / gross weight from the packing list.
  VF (INFOR NEXUS) INVOICE + PackingList .XLS (Timberland ...): description = goods description + style, HTS, COO, qty,
    price; net / gross weight per PO line from the PACKAGE DETAIL of the packing list (.xls needs xlrd - in the exe).
  Total Outer Pack UOM: PAT / CTN / BAG from the documents when the setting is the default PKG.

  SHELL EASTERN TRADING (lubricants export, e.g. to Shell Tianjin) - Tax Invoices + Delivery Notes + DHL/Danmar sea waybill
    Item value = Net Price - Second Distrib Cost (per line; the Invoice sheet total = sum of the items).
    Net weight + HS code from the DELIVERY NOTE (packing list) of each delivery (Tariff No / Total Net); a different invoice
    net weight is shown as MISMATCH. Delivery note not added -> invoice net weight + B/L HS, listed as CHECK. COO SG.
    Gross weight "with pallets" = delivery note / invoice gross + pallet weight (480 KG).
    One line per delivery (container); delivery notes are matched by delivery number - missing ones are listed.
  DHL E-MAIL INSTRUCTIONS pasted in "Cargo details" (export or import):
    CARRIER -> carrier agent, CARRIER BOOKING REF -> Outward MAWB when no master B/L, PLACE OF RELEASE (address -> O),
    PLACE OF RECEIPT, TYPE OF PERMIT, PERMIT REMARKS -> Trader's Remarks, SPECIAL INSTRUCTIONS:
      "SOB: 14 OCT" -> departure date, "INCLUDE DESCRIPTION: LUBRICATING OIL" -> added in front of every Item Description,
      "USE 6 DIGITS HS CODES" -> HS Code on Certificate = 6 digits (+ CO Certificate Description), HS Code stays 8 digits,
      "GROSS WEIGHT ... WITH PALLETS" -> the instruction gross weight is used.
    The DHL system block "LABEL:old/new" is read: HOUSE WAYBILL, MASTER WAYBILL, DISCHARGE PORT CODE, FLIGHT/VOYAGE,
      VESSEL NAME, TERM TYPE (Term Type on the Invoice sheet), TOTAL OUTER PACK + UOM (e.g. 10 UNT), TOTAL GROSS WEIGHT.
    These pasted values win over the documents for gross weight and outer packs (a CHECK note shows the difference).

  PANASONIC OPERATIONAL EXCELLENCE (Japan) "INVOICE & PACKING LIST" (import to Panasonic Factory Solutions Asia Pacific,
  e.g. KWE85298204-DOC.pdf) - usually a scan with the pages turned sideways: OCR turns every page the readable way up.
    Items: case no / part no / qty / unit (PIECE, UNIT, PACK, PAIR, CAN) / unit price / amount; description under the part no
    ("1PACK = 20PIECES" kept). qty x unit price = amount repairs OCR misreads (listed as CHECK).
    Net / gross weight per CASE (first item of the case; a case range 60298-60300 uses the range total) - a case with
    several items shares its weight over them by value. Checked against the grand total (57 cartons, net 140.600,
    gross 164.7, JPY total, PIECES / PACKS / UNITS ... totals).
    COO: "MADE IN ..." under the item, else Japan (CHECK note with the count). Brand: the maker printed under an item
    (KOGANEI, ORION KIKAI, PANASONIC) - "Unbranded for all" fills only the other items.
    No HS on the invoice: type it on the Item tab (tick "Save to product master" = next shipment fills by part no).
    HAWB from the file name (KWE85298204 -> 85298204, CHECK); Inward Transport Air + loading port JPNRT from
    "Shipped per AIR FREIGHT from TOKYO" (an air waybill added to the job wins). Incoterm FOB, currency JPY.
  OCR of turned scans (any supplier): a page read badly is tried at 90 / 270 / 180 degrees; the angle that works is used
    for the next pages (log: "pages are turned - reading them rotated 90 degrees").

14. JAPAN FOOD (GLT) - the JAPAN FOOD PERMIT TOOL inside the server   (New job > "Japan food (GLT)" or Auto)
   Documents of a job: GLT CO.,LTD invoice Excel + ANA / JAL air waybill PDF.
     Optional in the same job: the SFA "HS CODE & PRODUCT CODE" master list (RPT-xxxx.xlsx) and a previous permit Excel
     (its Header / Invoice / Brand are used: importer, UEN, agent, locations, licence ...).
   SFA master list: Settings > Product masters > Japan food (GLT) > Upload (checked - a wrong file is refused), or added to
     a job: kept in PERMIT_SERVER_DATA\masters\JAPAN_FOOD_SFA and used for every next job.
   Same rules as the desktop tool (permit_core.py unchanged): HS + product code only from the master list (exact item +
     prefecture, else same item of another prefecture = CHECK), net KG = invoice quantity, HS qty in the HS UOM, CASC qty =
     HS qty, description NAME**PREFECTURE**, Casc sheet = prefecture code, Shipping Marks 1 = group F / H / Z, MAWB /
     flight / arrival / pieces / gross weight / loading port / freight (weight charge + MYC / SCC ...) from the air waybill,
     Trader's Remarks "INVOICE NO : ...", the tool's OUTPUT FILE format.
   Item tab: type HS Code / CASC Product Code 1 (and tick "Save ... to the product master") = the tool's learned codes:
     the next invoice with that item + prefecture fills by itself. HS UOM / HS qty follow at once.
   One Excel per product group F / H / Z ("PARTS OF PERMIT"): Requirements > tick "split" (One Excel per ... item rows).
   HS UOM / control / CA product codes: the server's UPDATE HS CODE list (Settings) - one list for every tool.

15. BRAND (every tool)
   Brand from the documents / product master. When it is not there: Item tab > "Brand for items without one":
     "Unbranded for all"  -> Unbranded = Y on every item without a document brand;  or type a brand > "Apply to all".
     A brand read from the documents (e.g. the maker under a Panasonic item) always stays.
   The choice is remembered for that supplier (e.g. SHELL EASTERN TRADING = UNBRANDED): the next job of the supplier gets it
   automatically (log: "Brand: ... remembered"). "From documents" removes it again; Settings > Product masters lists the
   remembered brands (Forget).

16. IMPORT / EXPORT - each permit shows and fills only its own side
   The direction of a shipment comes from the documents (or Requirements > Direction = IMPORT / EXPORT).
   EXPORT permit: Exporter, Outward MAWB / transport / carrier agent, Out flight / transport identifier, Departure Date,
     Discharge port, Final destination, Consignee, End user, Certificate of origin (CO), Out HAWB / Out MAWB on items.
   IMPORT permit: Importer, Inward MAWB / transport / carrier agent, In flight / transport identifier, Arrival Date,
     In HAWB / In MAWB on items.
   Shared by both: declaration, loading port, release / receipt location, gross weight, outer packs, forwarder, remarks ...
   - Header / Item tabs, "Show every field", Requirements (job header values, More permit fields): only that side is
     shown, with the right names (e.g. "Departure date" for an export, "Arrival date" for an import).
   - The Excel / XML: a value of the other side is never written - not from a settings default, not from a value typed
     earlier, not from any tool (log: "... value(s) of the other direction removed").
   - Dates from the documents: export = departure / shipped-on-board date (ETD / ATD / SOB), import = arrival date
     (ETA / ATA). An export with only an ETA in the documents gets a CHECK: "type the departure / on-board date".
   EXCEL_ALL.zip now holds the final Excel files (brand, typed values) - same as the single downloads.

17. CUSTOMER RULES + AUTOMATIC READER
   - New job: drop the PDFs - there is no supplier list any more; the tool / reader is chosen from the document layout
     (again when PDFs are added to a job later).
   - Job > Requirements > Customer rules: after reading, the rules of the customers in the documents are listed, all
     ticked. Untick a rule you do not need for this job - the items are worked out again at once.
       SANDVIK: Export permit only | Description = TOOLS OR PARTS INV NO : ( invoice no ).
       DOW CHEMICAL PACIFIC: Export job | Model = shipment number | Brand = DOW | Description + Our Reference |
         Invoice number on the Invoice sheet (instead of "-").
   - Validate shows "Customer rules: n used, n unticked" (click Fix to open the tab).
   - Sea port names on a B/L (NHAVA SHEVA, SINGAPORE ...) give the port code (INNSA, SGSIN ...) - listed as CHECK.

18. TRAINED (built in): Sandvik TAX INVOICE (was a plugin - an old copy in PERMIT_SERVER_DATA\plugins is skipped),
   DOW Chemical Pacific export set, Vulcan Industrial Engineering import set. Invoices without any HAWB / B/L in the
   job: one permit per supplier + consignee. Sandvik export description = "TOOLS OR PARTS INV NO : ( invoice no )".
   Customer rules only follow the SELLER: an import consigned to Sandvik is a normal import (no Sandvik rules).

19. HS LISTS BY DESCRIPTION (per customer) - Settings > Product masters > "HS lists by description"
   Built in: PANASONIC FACTORY (HS_CODE_PANASONIC_FACTORY.xlsx, ~16,000 descriptions) - used for jobs whose documents name
   PANASONIC (e.g. the KWE / Panasonic Factory Solutions import).
   An item WITHOUT an HS on the invoice / in the product master gets:
     1. the HS of the SAME description (the most used code when the list has several), else
     2. the HS of the NEAREST description (same main word, most words in common) - listed as CHECK with the description used,
     3. nothing near enough -> HS stays EMPTY (type it; tick "Save to product master" to keep it for the item no).
   Order: invoice HS > product master > HS list by description. HS UOM / HS qty / CASC follow as usual.
   New version: "Upload new version" (the old file is kept as .bak) - used from the next Read again / Make files.
   Another customer: type the name + the name printed in its documents, "Add list (.xlsx)" (columns Description + HS code).
   The lists are in PERMIT_SERVER_DATA\masters\hs_desc and in every backup.

20. PERMIT SCREENS LIKE TRADENET (every job)
   Tabs: Header | Party | Cargo | Invoice | Item | CPC | Summary  (+ "Check & files" = validation, downloads, notes, log;
   Requirements = rules). Same sections as the TradeNet permit page; an IMPORT permit shows the import sections (Importer,
   Inward carrier agent, Inward transport ...), an EXPORT permit the export sections (Exporter, Outward carrier agent,
   Outward transport, Consignee, End user, Certificate of origin ...).
   - Header: job information (Settings > Declarant), declaration type (INPDEC / OUTDEC), cargo pack type, transport mode,
     BG indicator, reference documents, licences, certificate of origin (export).
   - Party: importer / exporter, carrier agent, freight forwarder, handling agent, consignee, end user, manufacturer.
   - Cargo: transport details (mode, date, ports, voyage, vessel, ocean B/L / MAWB), locations, container table.
   - Invoice: invoice information, supplier / manufacturer, invoice charge table (total, other taxable, freight, insurance).
   - Item: ONE item per page - Previous / Next / "1 of N" / go to; HS code, description (characters), COO, brand, model,
     HAWB / OBL (Copy to all), quantities + UOM, value, packing, CASC, CO (export); "View all items (table)",
     "View consolidation", "Item details ..." (where every value came from). A new HS code re-works HS UOM / HS qty / CASC.
   - CPC: the CASC / product code sheet.  - Summary: counts and sums, outer pack, gross weight, references, remarks.
   Yellow = filled in the Excel; typing saves into the Excel at once (the reset arrow puts the document value back).

21. INFINITY SOLUTIONS BRANDING + JOB LIST AS A DATA TABLE
    Logo in the side bar, on the log-in page, browser tab icon and the .exe icon (web\logo.png, web\mark.png,
    web\favicon.png, logo.ico). Colours: Infinity blue + orange.
    Jobs page: search + Status / Type (Import / Export) filters, "More filters" (supplier, only my jobs, show archived),
    Export (CSV for Excel), New job. Titles stay on top while scrolling, the tick box + JOB NO. columns stay on the left,
    click a title to sort, drag the edge of a title to make a column wider (remembered). Tick jobs -> Mark done /
    Archive bar. "..." on a row: open, open in a new tab, copy job no. / HAWB, mark done, archive.
    Footer: "Showing 1 to 25 of N results", rows per page, Previous / page numbers / Next.
    Item > "View all items (table)": the # column stays on the left, 50 rows per page (25 - 500), "Item form" goes back.
    Jobs page: the job route at the top (Reading > To check > Excel ready > Done > Permit approved, + Errors) shows
    how many jobs are at each stage - click a stage to see only those jobs (click again = all). The Jobs menu shows
    the number of jobs to check. Keys: "/" = search, "n" = new job. UPDATED column (e.g. "5 min ago").
    New job: drop the files or choose them (PDF / XLS / MAIL tags, remove with x), "Upload and read" at the bottom.
    Job page: "Jobs / name J3" at the top goes back to the list; empty job information shows "-".
    Permit screens show only the sections you use: Header = Declaration type + Declaration information;
    Party = Importer, Inward carrier agent, Freight forwarder (export: Exporter, Outward carrier agent, Freight forwarder);
    Cargo = Inward (export: Outward) transport details + Container details; Invoice = Invoice information,
    Supplier/Manufacturer party, Invoice charge; Item / CPC / Summary as before. "Show the other sections" at the
    bottom opens the rest when a job needs them (remembered per browser).
    Jobs list view: "Create job", Type, More filters, Export and the job filter box on one line; status filter
    boxes (count + status: Reading / To check / Excel ready / Done / Permit approved / Error) - tick one or more,
    "Clear status filters" shows all again, "x of y items" on the right; compact rows with a status icon.
22. CARGO / PORTS / TERM TYPE / ITEM ROWS
    Cargo: Location information first (always shown) - magnifier = pick PPZ / ME / KZ / JZ / TZ / CZ / O.
    Ports (Loading / Discharge / Next / Last): magnifier = search by code or name (e.g. CNASN or ANSHAN); typing a known
    code fills the name, a known name fills the code. Ports read from the documents or typed by you are added to the
    list automatically. Settings > Product masters > Port list: upload a full list (Excel: A = code, B = name).
    Selects show the TradeNet code too (1: Sea, 5: Other Non-Containerized ...) - lists from the dec-upload Excel.
    Export permits: Term Type is not filled automatically any more (import permits as before).
    Item page: "Consolidate" choice in the bar (No / same item + HS + COO / same HS + COO) - items worked out again.
23. TRAINED: Perkins UK "Customs Only Invoice" + DHL sea customs instructions + ONE arrival notice / sea waybill (see
    TRAINED_DOCUMENTS.txt 12b, 20b, 20c). Item page: "Copy to all" under every field (a qty + UOM pair is copied together).
24. RS COMPONENTS (was the RS COMPO TOOL) - built in: drop the consolidated invoice Excel + DHL customs instructions +
    B/L on New job. Customer rules tab: combine by HS + COO, CASC from the RS product code master, 50 items per Excel /
    IDAMISC / PCD separate, RS Trader's Remarks (type the exchange rate + J start there), invoice no. on the Invoice sheet.
    Settings > Product masters > "RS Components product code master": download / upload a new version (.xlsx or .xls).
    The FrtImport XML is made for every Excel as for all jobs.
25. SPLIT PERMITS / TNE / UEN
    A job that makes several Excel files (split permits): a Header / Invoice / Container value typed on one is written
    in ALL of them (not the Trader's Remarks and gross weight - they belong to each file). Item page "Copy to all" =
    every item of every Excel of the job. "Permit n of N" on top of the page switches the file.
    Ocean (Sea) permits: Total Gross Weight in TNE (KG / 1000, 3 decimals), UOM TNE. Air permits stay KGM.
    Container weight (TNE): round figure - 5.5 and above -> 6, below 5.5 -> 5 (at least 1).
    CR UEI: bizfile.gov.sg has a reCAPTCHA (no automatic look-up); the same ACRA register is open data on data.gov.sg.
    After reading, an importer / exporter without UEN is looked up there (exact company name, live) and remembered.
    Party page: magnifier next to CR UEI = search the ACRA register by name and pick. Needs internet on the server PC.
26. ONE VIEW, SAVE, THEN MAKE EXCEL / XML RULES
    The job pages show ALL items in one permit (no "Permit 1 of 10"); the split (RS 50 items, IDAMISC / PCD ...) is made
    only with "Make Excel + XML" - every value you typed goes into every split file (its own items).
    Changes are not saved at once: a bar at the bottom shows the summary (field, new value, how many items) - x = undo
    one, Discard = undo all, Save changes = write them. Make Excel saves open changes first (asks).
    FrtImport XML: ShptJobNo / ShptJobDate left blank. Items with HS UOM KGM / TNE (net weight / net weight based):
    ItemWgt + ItemWgtUnit filled, ItemQty / ItemUnit blank. Other UOMs: ItemQty + ItemUnit as before.
    Sea job with container(s): Total Outer Pack Qty = number of containers, UOM UNT (every tool).
27. HS BY DESCRIPTION WORD / TERM TYPE CHARGES / ROAD
    No HS on the invoice / product master / HS list by description: a word of the description gives the HS -
    CAPACITOR(S) = 85322400, INDUCTOR(S) = 85049090 (Settings > Rules > "HS by description word", WORD=HS per line).
    A CHECK note lists the items. An HS printed on the invoice always wins.
    Invoice Charge rows follow the Term Type (as TradeNet): CIF = Total Invoice + Other Taxable Charge only;
    CFR = + Insurance Charge (1 % SGD); CNI = + Freight Charge; EXW / FOB = + Freight + Insurance. Changing the
    Term Type on the Invoice page clears the charges that term does not have (saved with Save changes).
    Invoice readers can give the transport mode (Taiyo Yuden SG -> Malaysia = Road) when no B/L / air waybill has one.
28. HS ONLY FROM THE UPDATE HS CODE LIST / CUSTOMER PRODUCT LISTS
    An HS code is used only when it is in the UPDATE HS CODE list (or, when "Supplier HS -> Singapore code" is on, when
    its 6-digit subheading is in the list, e.g. EU 84099900 -> 84099979). A number read as HS that is not a real HS
    (e.g. product no. 14289560) is not used - HS left empty, MISSING note on Check & files. Order: invoice > product
    master > customer product list > HS list by description > description word. No HS anywhere = HS empty.
    Customer product lists (Settings > Product masters): e.g. the Qnity / DuPont DSP / EKC GMID master list (built in).
    Every sheet with Product ID / GMID / Material / Item No + HS Code (+ CA/SC Code) is read. The item no. of the
    invoice line (also "D14289560" printed apart as "D" + "14289560") finds the row: HS + CA/SC Code -> CASC Product
    Code 1 (NA / TBA / #N/A = no code). Upload more lists or a new version (same file name = replaced).
29. DUPONT DIVERSIFIED INDUSTRIALS (import) - BDP pack "1197631758 - Docs" + ONE arrival notice + DHL customs instructions
    Reader: Specialty Electronic Materials Switzerland "LEGAL I/C INVOICE" (one or more in the PDF, numbers like 4.267,66),
    the ONE sea waybill pages and the packing list in the same PDF. Per line: material, qty + unit, exact Net / Gross
    Weight of the line, COO, "Our Reference" (= Customer Order Number). Item charges on a line (FREIGHT 20,48 ...) are
    ADDED to the item value (4.267,66 + 20,48 = 4.288,14 = the line total on the invoice).
    Customer rules (tab Requirements > Customer rules), found from the importer name:
      - Vendor product list: choose the list / sheet (Settings > Product masters > Customer product lists) - the
        material number gives HS Code, CA/SC Code (CASC Product Code 1) and Material description. Auto = every list,
        No list = none. Materials not in the list are listed (MISSING) - add them to the list and upload it again.
      - Description = Customer Order Number : material number - description from the list (invoice text if missing).
    HS order: invoice > product master > vendor list > B/L "HS-NO" of the same net weight (only codes in UPDATE HS CODE).
30. MANY PEOPLE (e.g. 200 users)
    Users page > "Excel template" > one row per person (User name, Name, Password, Role) > "Add many from Excel".
    Existing user names are updated (name / role / new password), wrong rows are listed.
    PERMIT_SERVER_DATA\server_config.json:  "threads": 32 (people browsing at once, minimum 24)
                                           "workers": 2  (jobs read / Excel made at the same time; 4 on an 8-core PC)
    Server PC for 200 users: 8 cores, 16 GB RAM, SSD, wired network, never sleeps, UPS. Or a cloud server (Dockerfile).


31. TEAMS (DGF OCEAN / KWE ...) AND THE ITEM PAGE
  - New job: choose the Team (remembered). The customer rules (DuPont DDI vendor list + description + Casc sheet,
    Givaudan HS by material, VF brand + HS / origin / brand consolidation) and the customer product lists work only
    for the job's team. Panasonic HS list = KWE. Other teams / "Other": normal processing.
  - Settings: Teams (one per line) and, for every customer product list, the team that uses it.
  - Item page: Team, Other charges (invoice total only / include with the last item / Other Taxable Charge) and the
    customer rules (tick / untick, vendor product list) - no separate Requirements page any more; the shipment / item
    checks are under Check & files.
  - DuPont DDI (DGF OCEAN): Casc sheet = Item Number 1, 2, 3 ..., Product Code Seq Number 1, Code 1 = Control
    Substances of the DDIS master (blank for #N/A).
  - One common reader for every supplier ("Any supplier") - known layouts are listed in TRAINED_DOCUMENTS.txt.

32. DECLARATION ADVISOR (more intelligence)
  - Header tab, blue box: the permit type of the job and the declaration type the documents point to (Singapore
    Customs Module 5 rules), why, and what that type needs (lot no., previous permit no., IRAS registration ...).
    "Use ..." switches the permit type in one click.
  - Check & files shows "Declaration:" notes, e.g. dutiable items on an IPT GST permit (-> IPT DNG or bond under INP
    APS), dutiable and non-dutiable goods together on INP APS (split), liquor / tobacco in FTZ storage (INP SFZ),
    Place of Receipt ME / IGDS / TRADESP / DUMP ... not matching the permit, OUT DRT without an FTZ / exit point.
  - Uses the UPDATE HS CODE list (Duty items + Control columns) - upload a new list and the advisor follows it.

33. SEVERAL HOUSE B/Ls IN ONE PERMIT (e.g. Satair) / SEVERAL CURRENCIES
  - When the arrival notice / customs instructions list several house B/Ls under one ocean B/L, all of them become one
    permit (file = ocean B/L) and every item keeps its own In HAWB. Item rows: "Same HAWB + HS code + Origin (COO)".
  - Satair: upload the .zip as it comes (New job accepts .zip - the PDFs inside are used). Every line of every house
    B/L is read (e.g. 1,258 lines, each HBL total = its proforma total; all lines HAWB-wise in RAW_DATA.xlsx) and the
    items are consolidated by HAWB + HS code + COO by default (458 rows). Pick another "Item rows" on the Item page
    to change it.
  - Invoices in different currencies (e.g. Organon TWD / AUD / IDR / MYR): the Invoice sheet gets one invoice per
    currency (invoice numbers of that currency, total, insurance %), items are linked to it by Invoice Number, and the
    XML has one <Invoice> block per currency. Invoice page: buttons to switch between the invoices.

34. HS QTY BY THE SINGAPORE UNITS OF QUANTITY (HS UOM of the UPDATE HS CODE list)
  KGM kilogram = net weight | TNE tonne = net weight / 1000 | NMB piece = invoice pieces | TEN = pieces / 10 |
  CEN = pieces / 100 | MIL = pieces / 1000 | TPR (ten pairs) = pairs / 10 (invoice in PR / PAIR or pieces) |
  MTR metre (M, CM, MM, KM, FT, IN, YD converted) | LTR litre (L, ML, CL, M3, GAL converted; invoice in KG:
  LTR = KGM, 1 kg = 1 litre - invoice kg, else the net weight) | MTK square metre (M2, SQFT) | MTQ cubic metre (M3, CBM, L, CFT) |
  CTM carat (CT, G x 5) | GWH gigawatt-hour (MWH, KWH) | VAL = value in Singapore dollars (SGD invoices).
  Not convertible (e.g. LTR but the invoice is in KG) = HS Qty left blank with a CHECK note - type it on the Item page.

35. SUMMARY PAGE SPLIT / NAMES BY UEN / 2-LINE NAMES / CIF
  - Summary page: "Split the Excel every 50 item rows" (any number) + "keep one HS code in the same Excel" (an HS code
    with more rows than the maximum is the only one cut). Used when you click Make Excel + XML.
  - Names by UEN: the importer / exporter (+ exporter address), freight forwarder and carrier agent names are taken
    from the ACRA register (data.gov.sg open data) for their UEN - automatically after reading (needs internet on the
    server PC; remembered), or with "Name from UEN" on the Party page. Settings > options to switch it off.
  - Party page: importer / exporter name in 2 lines of 35 characters, as on the TradeNet permit.
  - Invoice page: Cost, Insurance and Freight worked out from the invoice + the charges of the Term Type (CIF: none,
    CFR: insurance, CNI: freight, FOB / FCA / EXW: both); type another value to change it ("auto" puts it back).

36. SEVERAL INVOICES / PACKING LISTS IN ONE PDF
  - Upload a combined PDF as it is: it is split per document (each "Page 1 of N") and all invoices + packing lists
    are read and matched (CommScope all sites, Maquet, Organon, Givaudan, Arnold Suhr, Allison, Halliburton,
    SharkNinja). The log shows "N documents in one PDF". Check the invoice count on the Invoice page.

20. SFA TEAM (Japanese perishables - SFA rules)   New job > Team = SFA  (Settings > Teams lists it on every server)
   Japan food (GLT) jobs of the SFA team use the SFA project rules instead of the plain Japan food rules:
   - HS + CASC code: codes saved on the Item tab > SFA overrides > SFA MASTER DATA > past permits (RPT reports, valid in
     UPDATE HS CODE) > same Japanese name in past permits > nearest description (CHECK). No RPT list needed in the job.
   - Net weight = quantity x standard (HOTATE 12 PC x 150 G/PC = 1.8 kg), KG lines = quantity.
     Importers listed in Settings > SFA team > "Importers: invoice (Quantity) = net weight" (GINZA SUSHI ICHI PTE. LTD.):
     the invoice (Quantity) column is the net weight of every line.
   - HS qty by HS UOM, CASC qty by PRODUCT UOM: NMB = qty, KGM = net weight, TNE = net weight / 1000, TEN = qty / 10.
   - Casc sheet only for items with a CASC code (non-controlled items such as packing paper have none).
   - Every code + source is shown on the Item tab (HS from / CASC from) and every CHECK in Notes.
   Settings > Product masters > "SFA team": counts, files, download the overrides .csv / rules; admin: upload a new
   SFA MASTER DATA or RPT report (.xlsx / .xls - the database is rebuilt at once), upload the overrides .csv, Rebuild,
   Remove a report. Data: PERMIT_SERVER_DATA\masters\sfa (in every backup). Rules: masters\sfa\SFA_RULES.md.
   Other teams: Japan food jobs work exactly as before.

21. AGILENT "IC CUSTOMS INVOICE" (DGF OCEAN team, trained 09/10/2026)   New job > drop the IC invoices + packing lists (Auto = Agilent)
   Agilent intercompany IC customs invoices (e.g. Taiwan -> SG10 / SG1U, CA HAND CARRIES, Carrier Ref CR9-10, MEA#1793):
   - every line read (2-line descriptions, CoO, National HTS -> 8 digits, ECCN), totals checked (Amount in USD / Total weight).
   - weights: Box Weight is the whole box on every line - counted once per box, split by value, NW = 85 %.
   - no HAWB on the documents: all IC invoices of the same supplier + consignee = ONE permit; paste the HBL / OBL / vessel /
     voyage / ETA in Cargo details (or type them on the Shipments / Header tab).
   - importer = Sold-To (AGILENT TECHNOLOGIES SINGAPORE (INTERNATIONAL) PTE LTD), UEN from "RN:" (201815040D).

22. XML MESSAGE TYPE PER JOB + CARRIER AGENT BOOK (09/10/2026)
   XML message type (FrtImport ShptMsgType):
   - Job > Header and cargo > "XML message type": pick or type the type this customs job needs (list = Settings).
   - Blank = Settings: "Message type rules" (TEAM or NAME=TYPE: a job of that team, or whose importer / exporter name
     contains NAME, gets that type), else "Message type import / export". Order: job choice > rule > default.
   - Settings > XML: "XML message types (one per line)" = the list offered on every job. The log shows the type used.
   Carrier agent book (Settings > Rules > "Carrier agents (UEN=NAME, one per line)"):
   - Every permit (inward + outward): carrier agent UEN filled -> name copied from the book (book name wins);
     only the name filled -> UEN from the book. Option "Carrier agent name copied by UEN" (on by default).
   - Job > Header and cargo: "Carrier agent (book)" pick list fills name + UEN; typing a UEN fills the name.
   - SATS (197201770G = SINGAPORE AIRPORT TERMINAL SERVICES) is in the book by default - add your carriers.

37. DECLARATION ADVISOR DECIDES (Header tab)
  - The advisor (Singapore Customs Module 5 rules) now FILLS the permit: Declaration Type + code, Release Location O for
    exports, Receipt ME / DUMP when blank, the Message Type (IPTDEC / INPDEC / OUTDEC) and the XML message type (when that
    type is in Settings > XML message types).
  - Who wins: your choice (Declaration type picked on the Header tab or permit type picked on the job) > customs
    instructions (kept, but corrected automatically when the rules say it cannot be right, e.g. Place of Receipt of
    another type) > the advisor. A value typed in a Header cell is always written last, so it always wins.
  - Header tab card: "applied: ..." and who decided; "Declaration type" list to change it, "Auto - advisor decides" to
    give it back; untick "apply the advisor automatically" = advice only (nothing filled).
  - A type the upload template does not have (e.g. IPT DNG for dutiable goods, INP DES) is never guessed: ERROR note
    "declare it in TradeNet directly or choose another type".

38. DOCUMENT ANALYSIS (Check & files > Document analysis; RAW_DATA.xlsx sheets DOCUMENT_ANALYSIS + DOCUMENT_CHECKS)
  - Every document: type + reader, pages (combined PDFs), the values read (invoice no., currency, total, weights,
    HAWB ...) and what was NOT found, line completeness (% of lines with description, qty, amount, HS, COO, weights),
    an extraction score 0-100, and its own checks: line amounts = invoice total (charges / goods value explained),
    qty x unit price = amount (per 10 / 100 / 1000 prices allowed), zero / negative lines, NW > GW, PL line weights =
    PL totals, OCR read, not recognised, not matched to a shipment.
  - Documents compared per shipment: invoice totals = permit items (per currency, charges), duplicate invoice
    numbers, invoice <-> packing list (references, items found, qty per item), gross weight and packages from every
    document (packing list, invoice, B/L - several house B/Ls added up, customs instructions, arrival notice),
    net weight > gross weight, several HAWB numbers, consignee / shipper name vs the importer / exporter, HS code
    changed from the invoice, items without HS / COO, estimated net weights.
  - "Where the permit values came from": count of items per source of HS code, origin, net weight, HS UOM, CASC.
  - The log shows the summary after every Read again.

39. LIQUOR (WINE / CHAMPAGNE / SPIRITS) - CLASQUIN TEAM, PERMIT TYPE "ID" (IPT DNG - Duty & GST)
  - New permit type ID on the job = IPT DNG: Declaration Type "Duty and GST" / DNG (Excel), ShptDeclType DNG (XML).
  - Singapore HS + CA/SC product code by the liquor rules (liquor_rules.py) from type, colour, bottle size and alcohol %,
    checked against the official product code descriptions of the UPDATE HS CODE list. Supplier (EU) HS codes are never
    copied. When the rules cannot decide (whisky blended or straight unknown, wine colour missing, unknown spirit) the
    code stays blank with a CHECK - type it on the Item page.
  - Each wine / spirit is its own item row (never merged with another strength or bottle size).
  - Update the Clasquin past-permit list: replace masters/CLASQUIN_MASTER_LIST.xlsx (HS Code | Description Goods |
    CA/SC Product Code 1) - it is used for a cross-check note only.

40. NEW LOOK + DASHBOARD (Infinity Solutions brand)
  - Home page = Dashboard: brand banner (jobs this month, Excel + XML made, permits approved), tiles New today / To
    check / Excel ready / Awaiting permit no. / Errors (click = the job list filtered on it), jobs created in the last
    14 days (hover a bar; "Show as table"), pipeline by status, jobs by supplier reader, import / export, most active
    users, recent jobs.
  - Slim icon menu on the left (hover = name), search box at the top (Enter = search the jobs), your name at the top
    right (change password / log out). Jobs list: status tabs, filters, search, tidy table with status pills.
  - Files: web/theme.css (look) and web/dashboard.js (dashboard) - the job pages and TradeNet tabs work as before.

41. PERMIT TYPE DECIDES IMPORT / EXPORT
  - OD / OT = export permit, IG / ME / IT / IM / ID = import permit: every tab (Header, Party, Cargo, Item) then shows
    only that direction's fields, also when the documents look like the other direction (MISMATCH note).
  - In that case the documents' transport (ports, vessel / flight, date, OBL / MAWB, containers, release / receipt)
    belongs to the other leg and is left blank with a MISSING note - type the outward (or inward) details.
  - Legrand invoices: lines in metres ("420.000 M") are read too.
