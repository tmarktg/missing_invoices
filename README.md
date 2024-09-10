**Purpose**

- This code is meant to find the missing PO\#s needed to record item part prices in the excel sheet to update prices to market trends
- The problem my boss had was that there was always missing invoice numbers, so after I fixed the invoices, there would always be missing invoices, so he would have to dig into the database, find the missing PO\#s and add it to the PO\#s list
- Now the program finds the missing PO\#s list by using pandas and set manipulation

**How it works**

- Since the data is confidential here is dummy data to show

**File Directory:**

**Desktop**

- PO\#s
- excel_files

**PO\#s**  
![](/img/img-1.png)

**excel_files**  
![](/img/img-2.png)

- We check each invoice and take the purchase order numbers in order to see what purchase order numbers are missing from PO\#s

**example.xlsx**
![](/img/img-3.png)

- If there is a purchase order number in the example.xlsx which is (98) the program reports it and adds it to a set so that there is a list of missing invoices numbers

**Working version of the code**  
![](/img/img-4.png)
# missing_invoices
