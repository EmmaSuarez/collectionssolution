# collectionssolution
A collections solution system that i created, that uses the MS ecosystem, and allows users to log in accepted payments.

Users receive payments in both physical and digital form.
This solution allows users to report pyments to another sector, giving details such as the clients ID number, tax identification number, and the clients name.
The user gets this information by typing only the client ID number, and is retrieved using from another excel file using the VLOOKUP formula.
Since the users accept payments for two different companies, i devised a solution, using a cell with data validation and a list with both company names and the following IF formula:

=IF(B3="COMPANY1",+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY1'!$A:$C,2,FALSE),""),+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY2'!$A:$C,2,FALSE),""))

This formula displays the clients information based on what company is selected in cell B3, looking it up in different sheet within the clientsdata.xlsx file.

The user then proceeds to paste or type in the invoice number associated with this payment into cell D4.

Since the company is based in Argentina, we take payments in the local currency, which is Pesos. Almost all of the inventory sold is valued in US dollar, so the user now needs to enter the foreign exchange rate in which this payment will be accounted in cell D6.

In cell D8, the user inputs the total payment value, while other relevant transaction details, such as withholding taxs, are to be entered below.
Within ranges D10 and D14, users enter the values corresponding to several witholding taxes, that are available to be selected within a list in range B10 to B14.
Withholding taxes values are totalized in cell D15.

Cell B8 diplays the total of the payment in USD using the formula =(+D8+D15)/D6(Total payments plus total withholding taxes, divided by foreing exchange rate). 
