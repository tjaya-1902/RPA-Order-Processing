In this RPA project, a bot processes two types of emails from an inbox: order emails and restock emails.

Project Files:
- ProductStock.xlsx – Tracks the current available stock of each product.
- SalesOrder.xlsx – Stores details of processed sales orders.
- RestockHistory.xlsx – Logs restock events from restock emails.
- InvoiceTemplate.doc – A Word document containing placeholders for generating invoices.

Process Overview:

The bot first classifies emails by analyzing their subject line and stores them in two separate lists: one for order emails and one for restock emails.

Step 1: Process Restock Emails
- Extract restock details (date, product ID, and restock quantity) from the email body.
- Update the stock quantity of the relevant product in ProductStock.xlsx.
- Log the restock entry in RestockHistory.xlsx.

Step 2: Process Order Emails
- Extract order details (customer name, date, product ID, order quantity, etc.) from the email body.
- Compare the order quantity against the available stock in ProductStock.xlsx.
- If sufficient stock is available:
    - Subtract the ordered quantity from the current stock.
    - Record the order in SalesOrder.xlsx.
    - Generate a PDF invoice using InvoiceTemplate.doc and save it to the Invoices folder.
    - Generate an order confirmation email (with the invoice attached) and save it to the EmailsOutgoing folder.
- If stock is insufficient:
    - Generate and save a cancellation email to the EmailsOutgoing folder.
