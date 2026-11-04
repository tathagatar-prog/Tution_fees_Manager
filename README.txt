# Tuition Centre Manager

A mobile-friendly Streamlit app for **Tutoring Computer Science & Application**.

## Included
- Existing student register imported from `Tution Payment Status.xlsx` on first run.
- Add, edit, search, deactivate and permanently delete students.
- Month-wise paid / partial / unpaid fee status.
- Record full or partial fee payments.
- Compact A5 printable fee receipt with your supplied logo.
- Income and expense register.
- Payment history, CSV exports and SQLite database backup.

## Run on Windows
1. Install Python 3.10 or newer.
2. Extract this ZIP into a folder.
3. Open Command Prompt in that folder.
4. Run:

   ```bash
   py -m pip install -r requirements.txt
   py -m streamlit run app.py
   ```

5. Streamlit opens the app in your browser. To use it from a phone on the same Wi-Fi, find your computer's local IP and run:

   ```bash
   py -m streamlit run app.py --server.address 0.0.0.0
   ```

   Then open `http://YOUR-COMPUTER-IP:8501` on your phone. Allow Python/port 8501 through Windows Firewall if prompted.

## Important data notes
- The app imports the included workbook only when the SQLite database is empty and the first-import marker is absent.
- The workbook's filled date cells are retained as payment records. The fee month is taken from the month heading and year from the entered date.
- If a student's fee amount is blank in Excel, the payment date is retained with amount ₹0 and a note asking you to verify the amount. Please edit that student's monthly fee before recording new payments.
- Historical imported payments use payment mode `Imported` and receipt IDs beginning `IMP-`; they do not generate retroactive printed receipts.
- A repeated month/date entry in the workbook is preserved as a separate payment record rather than silently discarded.
- Deactivation preserves student and payment history. Permanent deletion removes that student's payment history and should be used carefully.
- The app is local-first. The database file `tuition_centre.db` is created beside `app.py`. Back it up regularly using the Backup & Data page.

## Receipt printing
After recording a payment, the receipt preview opens. Use the browser print dialog and choose your printer. The receipt is styled for A5 paper / compact printing; printer settings may allow further scaling.

## Mobile use
The interface is responsive. For phone access, the computer and phone should be on the same Wi-Fi network. Public internet access requires hosting/deployment and appropriate security; do not expose the local app directly to the internet.
