# Order Logger

A single-page tool for logging online order deliveries into an Excel sheet.

- **New delivery:** paste the "TODAY OUT FOR DELIVERY" WhatsApp message and the form fills itself (name, model, tracking, quantity, full mobile number from the last 4 digits, address from the pin code). Fill the rest from remembered dropdowns and add it to Excel.
- **Out for delivery:** every order marked "Out for Delivery", with a one-click Delivered button.
- **Payments:** log payments received.
- **Chat:** paste `#ORDER`, `#PAYMENT` or `#UPDATE` messages.

Everything runs in your browser. Your Excel file is read and saved on your own computer and is never uploaded anywhere. Use Chrome or Edge on a computer to save straight into the file.

## Google Sheets mode (phone, PC can be off)

`apps-script/Code.gs` is a small Apps Script web app that lets the page read and write your Google Sheet. Paste it into the sheet's Extensions > Apps Script, set your own `ACCESS_KEY`, and deploy it as a web app (Execute as: Me, Who has access: Anyone). Then press **Connect** on the page and enter the web app URL and key. The key is never stored in this repository.
