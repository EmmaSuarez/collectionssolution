# collectionssolution
I have created a collections solution system that utilizes the MS ecosystem, allowing users to log accepted payments.

Users receive payments in both physical and digital forms. This solution enables users to report payments to another department by providing details such as the client's ID number, tax identification number, and client's name. To retrieve this information, the user only needs to input the client ID number, which is fetched from another Excel file using the VLOOKUP formula.

Since users accept payments for two different companies, I devised a solution utilizing a cell with data validation and a list containing both company names. The following IF formula is used:

=IF(B3="COMPANY1",+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY1'!$A:$C,2,FALSE),""),+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY2'!$A:$C,2,FALSE),""))

This formula displays the client's information based on the selected company in cell B3, retrieving the data from different sheets within the clientsdata.xlsx file.

Next, the user pastes or types the invoice number associated with the payment into cell D4.

Since the company is based in Argentina, payments are accepted in the local currency, which is Pesos. However, most inventory items are valued in US dollars. Therefore, the user needs to enter the foreign exchange rate in which this payment will be accounted for in cell D6.

In cell D8, the user inputs the total payment value, while other relevant transaction details, such as withholding taxes, are entered below. Ranges D10 to D14 are used for users to enter values corresponding to several withholding taxes, which can be selected from a list in range B10 to B14. The total withholding taxes are calculated and displayed in cell D15.

Cell B8 displays the total payment amount in USD using the formula =(+D8+D15)/D6 (Total payments plus total withholding taxes divided by the foreign exchange rate).

Cell F6 is a textbox that enables users to enter additional information. This textbox interacts with the tasks created in Oulook when the user marks specific checkboxes below.

Cell B17 is an empty space designated for the user to paste a clipboard image, which can then be exported to the final PDF this solution creates.

Cells B36, D36, and F36 contain buttons that, when clicked, erase the information in the corresponding cell directly below them.

