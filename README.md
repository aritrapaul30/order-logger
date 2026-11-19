# Digital Den Order Logger

A phone-friendly web app for logging online orders, deliveries and payments into Google Sheets.

- **Dealer picker** and a **year picker**. Each dealer has one sheet per year.
- **Orders:** save an order when you place it (status Ordered). When the out-for-delivery message comes, tap the saved row (tick several rows for a multi-phone package) and paste the message to complete it.
- **Out For Delivery:** one tap marks a package Delivered and adds Amount + Commission cut to the client whose card paid.
- **Dealer Payments:** what each dealer still owes (delivered orders only), payments received, and **Start next year**, which creates next year's sheets and carries LEFT over as "Previous Year Due".
- **Clients Payments:** what you owe each client, Pay (cash or "settled via dealer cut"), Share statement, Mark settled.
- **Profit** (delivered orders) and **Tracking Search** across all dealers.

## Google Drive layout

```
Digital Den/
  Payment Details
  2026/  Rahul Yadav 2026, Sourav Mondal 2026
  2027/  ...
```

`apps-script/Code.gs` is the Apps Script web app that reads and writes these sheets. Its access key is set only inside your own Apps Script project, never in this repository.
