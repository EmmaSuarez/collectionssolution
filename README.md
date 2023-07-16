# Collections solution
I have created a collections solution system that utilizes the Microsoft ecosystem, allowing users to log accepted payments.

Users receive payments in both physical and digital forms. This solution enables users to report payments to another department by providing details such as the client's ID number, tax identification number, and client's name. The required information is fetched from another Excel file using the VLOOKUP formula when the user inputs the client ID number.

Since users accept payments for two different companies, I devised a solution utilizing a cell with data validation and a list containing both company names. The following IF formula is used:

=IF(B3="COMPANY1",+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY1'!$A:$C,2,FALSE),""),+IFERROR(+VLOOKUP($D$2,'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY2'!$A:$C,2,FALSE),""))

This formula displays the client's information based on the selected company in cell B3, retrieving the data from different sheets within the clientsdata.xlsx file.

Next, the user pastes or types the invoice number associated with the payment into cell D4.

Since the company is based in Argentina, payments are accepted in the local currency, which is Pesos. However, most inventory items are valued in US dollars. Therefore, the user needs to enter the foreign exchange rate in which this payment will be accounted for in cell D6.

In cell D8, the user inputs the total payment value, while other relevant transaction details, such as withholding taxes, are entered below. Ranges D10 to D14 are used for users to enter values corresponding to several withholding taxes, which can be selected from a list in range B10 to B14. The total withholding taxes are calculated and displayed in cell D15.

Cell B8 displays the total payment amount in USD using the formula =(+D8+D15)/D6 (Total payments plus total withholding taxes divided by the foreign exchange rate).

Cell F6 is a textbox that enables users to enter additional information. This textbox interacts with the tasks created in Oulook when the user marks specific checkboxes below.

Cell B17 is an empty space designated for the user to paste a clipboard image, which can then be exported to the final PDF this solution creates.

Cells B36, D36, and F36 contain buttons that, when clicked, erase the information in the corresponding cell directly below them. These cells are used by the user, to enumerate the number of payments that contribute to the total amount paid.

There are four checkboxes that, when selected by the user, generate different tasks in Outlook based on the information entered. These tasks are created when the PDF button is pressed.

Finally, when the user presses the PDF button, the program runs and creates a folder within the "Payments" folder in OneDrive.

-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Solución de cobros 
He creado un sistema de solución de cobros que utiliza el ecosistema de Microsoft, que permite a los usuarios registrar pagos aceptados.

Los usuarios reciben pagos en formas físicas y digitales. Esta solución permite a los usuarios informar los pagos a otro departamento al proporcionar detalles como el número de identificación del cliente, número de identificación tributaria y nombre del cliente. La información requerida se obtiene de otro archivo de Excel utilizando la fórmula BUSCARV(VLOOKUP) cuando el usuario ingresa el número de identificación del cliente.

Dado que los usuarios aceptan pagos para dos compañías diferentes, diseñé una solución utilizando una celda con validación de datos y una lista que contiene los nombres de ambas compañías. Se utiliza la siguiente fórmula SI(IF):

=SI(B3="COMPANY1";+SI.ERROR(+BUSCARV($D$2;'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY1'!$A:$C;2;FALSO);"");+SI.ERROR(+BUSCARV($D$2;'C:\Users\emmanuel_suarez\Desktop\[clientsdata.xlsx]COMPANY2'!$A:$C;2;FALSO);""))

Esta fórmula muestra la información del cliente según la compañía seleccionada en la celda B3, recuperando los datos de diferentes hojas dentro del archivo clientsdata.xlsx.

A continuación, el usuario pega o escribe el número de factura asociado al pago en la celda D4.

Dado que la empresa se encuentra en Argentina, se aceptan pagos en la moneda local, que es el Peso. Sin embargo, la mayoría de los productos comercializados tienen un valor en moneda dólares estadounidense. Por lo tanto, el usuario debe ingresar el tipo de cambio en el cual se contabilizará este pago en la celda D6.

En la celda D8, el usuario ingresa el valor total del pago, mientras que otros detalles relevantes de la transacción, como los impuestos retenidos, se ingresan a continuación. Los rangos D10 a D14 se utilizan para que los usuarios ingresen los valores correspondientes a varios impuestos retenidos, que se pueden seleccionar de una lista en el rango B10 a B14. Los impuestos retenidos totales se calculan y se muestran en la celda D15.

La celda B8 muestra el monto total del pago en USD utilizando la fórmula =(+D8+D15)/D6 (pagos totales más impuestos retenidos totales, dividido por el tipo de cambio).

La celda F6 es un cuadro de texto que permite a los usuarios ingresar información adicional. Este cuadro de texto interactúa con las tareas creadas en Outlook cuando el usuario marca una o más casillas de verificación que se encuentran debajo.

La celda B17 es un espacio vacío designado para que el usuario pegue(CTRL+V) una imagen del portapapeles, que luego se exportara al PDF final que esta solución crea.

Las celdas B36, D36 y F36 contienen botones que, al hacer clic en ellos, borran la información en la celda correspondiente directamente debajo. Estas ultimas celdas se usan por el usuario, para enumerar la cantidad de pagos que componen el total abonado.

Hay cuatro casillas de verificación que, cuando el usuario las selecciona, generan diferentes tareas en Outlook según la información ingresada. Estas tareas se crean cuando se presiona el botón de PDF.

Finalmente, cuando el usuario presiona el botón de PDF, el programa se ejecuta y crea una carpeta dentro de la carpeta "Payments" en OneDrive.

