[Brazilian Portuguese version](README.pt-br.md)

Solution developed in **VBA** to **automate the updating of a payment control spreadsheet** for a real state business.

## How it works
The solution works using at least two spreadsheets:
- the control spreadsheet (the one that will receive the data);
- and the spreadsheet that contains the receipt data (our data source).

When opening the control spreadsheet, at the first column and first line of the **Customers** tab there will be the **Update** button, which when activated, will ask for the file containing the payment data (in the demonstration, we use the file *RECEBIMENTO OUTUBRO DE 2022.xlsx*).

Once the file is selected, the algorithm will identify the necessary information by column names, such as apartment number, payment date and amount paid.
This information is then transposed to the appropriate fields in the **Customers** tab itself.

At the end of the transposition, a new tab will be created containing the data collected and the status of the transposition. If there is any record from the source spreadsheet that was not transposed to the control spreadsheet, the status will indicate *'Unit not found'* or *'Competence not found'*.