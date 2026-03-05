# auto_doc_product

**GR:** Παραγωγή τελικών εγγράφων Word (DOCX) από templates με placeholders, με ένα κλικ μέσω VBA macro.  
**EN:** One-click generation of final Word documents (DOCX) from placeholder templates using a VBA macro.

---

## Ελληνικά (GR)

### Τι κάνει
- Διαβάζει τιμές από ένα **Controller Word (.docm)** που περιέχει έναν πίνακα **ΠΕΔΙΟ / ΤΙΜΗ**
- Παίρνει όλα τα templates που ταιριάζουν στο μοτίβο **`TEMPLATE_*.docx`**
- Δημιουργεί φάκελο **`OUTPUT`**
- Παράγει **ξεχωριστά τελικά έγγραφα** στο `OUTPUT` χωρίς να πειράζει τα templates
- Αντικαθιστά placeholders τύπου `{{Key}}` με τις αντίστοιχες τιμές
- Προαιρετικά εφαρμόζει **κανόνες χρόνου** (ώρα έναρξης/περάτωσης) ανά έγγραφο:
  - 10’ διάρκεια για όλα
  - 20’ διάρκεια σε ένα συγκεκριμένο template-τύπο (π.χ. ανιχνεύεται από λέξεις στο όνομα αρχείου)
  - διάλειμμα `BreakMinutes` (2’/5’ κτλ) μεταξύ εγγράφων

### Πεδίο χρήσης
Το project **δεν είναι αποκλειστικά για χρήση από αστυνομικές/δημόσιες αρχές**.  
Είναι ένα γενικό εργαλείο “template → output” για Word, που μπορεί να χρησιμοποιηθεί με τον ίδιο τρόπο για **οποιοδήποτε έγγραφο** (π.χ. αναφορές, πρακτικά, αιτήσεις, βεβαιώσεις, εταιρικά έντυπα, checklists), αρκεί να υπάρχουν placeholders τύπου `{{...}}` μέσα στα templates.

### Σειρά εκθέσεων (πολύ σημαντικό για τις ώρες)
Για να υπολογίζονται σωστά οι ώρες έναρξης/περάτωσης, τα templates πρέπει να έχουν **σαφή σειρά**.

Προτείνεται να βάζετε **αρίθμηση στην αρχή του ονόματος αρχείου**, π.χ.:
- `TEMPLATE_1. ... .docx`
- `TEMPLATE_2. ... .docx`
- `TEMPLATE_10. ... .docx`

✅ Το macro ταξινομεί τα templates **αριθμητικά** (με βάση τον **πρώτο αριθμό** στο όνομα αρχείου), ώστε η σειρά να είναι πάντα προβλέψιμη και οι υπολογισμοί ώρας να βγαίνουν σωστά.

### Γρήγορο Setup (Word 2016 Windows)
1. Ανοίξτε το `examples/00_Controller_example.docx` και κάντε **Save As → .docm**.
2. Πατήστε **Alt+F11 → File → Import File…** και κάντε import το `src/ControllerModule_TimeOutput_Sorted_GR_ANSI.bas` (ή το αντίστοιχο `.bas` που χρησιμοποιείτε). Μετά κάντε αποθήκευση **Ctrl+S**.
3. Βάλτε τα δικά σας templates στον ίδιο φάκελο με το `.docm` και ονομάστε τα `TEMPLATE_*.docx`.
4. Βεβαιωθείτε ότι στα templates έχετε:
   - `{{OraEnarxis}}` στην αρχή (ώρα έναρξης)
   - `{{OraPeratosis}}` στο τέλος (ώρα περάτωσης)
5. Τρέξτε: **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run** είτε πιο γρήγορα:
- Πατήστε **Alt + F8**
- Επιλέξτε `Generate_Reports_To_OUTPUT_From_Table`
- Πατήστε **Run / Εκτέλεση**

> Αν τα macros είναι blocked: κάντε Unblock στο .docm ή προσθέστε τον φάκελο σε Trusted Locations.

### Σημείωση ασφάλειας
Το αποθετήριο περιέχει **απλό VBA κώδικα** και παραδείγματα templates.  
Δεν κάνει λήψεις/εκτέλεση εξωτερικού κώδικα, δεν κάνει δικτυακές κλήσεις και δεν επιχειρεί να “παραμείνει” στο σύστημα.  
Ο κώδικας είναι **πλήρως αναγνώσιμος/ελέγξιμος**. Όπως με κάθε macro-enabled λύση, παρακαλούμε να το τρέχετε μόνο σε αξιόπιστο περιβάλλον και να κάνετε έλεγχο με τα εργαλεία ασφαλείας που χρησιμοποιείτε.

---

## English (EN)

### What it does
**autodocproduct** is a lightweight Word automation workflow that generates finalized documents from reusable templates — and it can be used for **any document type**, not tied to a specific domain or organization.

- You maintain one **Controller** document (a `.docm`) containing a simple **KEY / VALUE** table.
- You create one or more **template `.docx` files** containing placeholders like `{{Key}}`.
- With a single macro run, the tool:
  - Reads the values from the Controller table
  - Copies each template into an `OUTPUT` folder (templates stay unchanged)
  - Replaces all placeholders with the corresponding values across the document (including headers/footers and text boxes when present)
  - Optionally applies time-based rules (e.g., calculated start/end timestamps per generated document)

Each run produces a clean set of output documents ready to share, print, or archive.

### Document order (important for time calculations)
To calculate start/end times correctly, templates should have a **clear and predictable order**.

We recommend adding a **leading number** to each template filename, for example:
- `TEMPLATE_1. ... .docx`
- `TEMPLATE_2. ... .docx`
- `TEMPLATE_10. ... .docx`

✅ The macro sorts templates **numerically** (based on the **first number** in the filename) to ensure a consistent, predictable sequence and correct time progression.

### Quick Setup (Word 2016 Windows)
1. Open `examples/00_Controller_example.docx` and **Save As → .docm**.
2. Press **Alt+F11 → File → Import File…** and import `src/ControllerModule_TimeOutput_Sorted_GR_ANSI.bas` (or the `.bas` you use). Then save with **Ctrl+S**.
3. Put your templates next to the `.docm` and name them `TEMPLATE_*.docx`.
4. Ensure templates contain:
   - `{{OraEnarxis}}` for start time
   - `{{OraPeratosis}}` for end time
5. Run: **Developer → Macros → Generate_Reports_To_OUTPUT_From_Table → Run** or quicly:
- Press **Alt + F8**
- Select `Generate_Reports_To_OUTPUT_From_Table`
- Click **Run**

### Scope of use
This project is **not exclusively intended for law enforcement / public authorities**.  
It is a general-purpose “template → output” Word automation tool and can be used the same way to generate **any kind of document** (e.g., reports, minutes, requests, certificates, corporate forms, checklists), as long as the templates contain `{{...}}` placeholders.

### Security note
This repository contains **plain VBA source code** and example templates.  
It does **not** download or execute external code, does **not** make network requests, and does not try to **persist** on the system.  
The code is **fully readable/auditable** in this repo. As with any macro-enabled workflow, please run it only in trusted environments and feel free to scan the files with your preferred security tools.

---

## License
MIT (see `LICENSE`).
