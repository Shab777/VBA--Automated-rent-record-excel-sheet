# VBA-Automated-rent-record-excel-sheet
Under this project I have developed and implemented an Excel VBA to streamline the administrative processes for letting agency company.
It optimizes the work speed and one can keep track of tenant's payment record for each month. 
There are two excel files. One stores the complete data of the company name as 'All Property Details' such as property addresses, Landlord names and tenant names. The second file is 'Tenant Payment Sheet' which keeps record of tenant's payments.
How does it work-
Write the date under date column when tenant makes the payment. In the next cell write 'FULL' if tenant has paid the amount or write 'amount in number' if tenant paid some amount of money. Leave this cell selected and click the check box. A message box will appear stating whether the tenant paid full amount or balance dues. It fill the entire row with different colors only to highlight for balances.
If any new tenancy comes with the existing property, make the neccessary changes and select the entire or the first cell of that row. Then go to 'My Macros' in taskbar, look for 'Tenant payments' and select 'New Changes'. It helps to keep a record of new changes.
It any tenancy ends, select that row, go to 'My Macros' in taskbar, look for 'Tenant payments' and select 'Move Out'. It helps to keep a track of ending tenancy.
If any new property comes, look for 'Property details' in 'My macros'. Select propety address, landlord names and tenant names through combo box. Which extracts data from 'All property Detials' folder.  
At the end of month, add a new sheet, copy all data from previous sheet and past it into new sheet. Rename with exist month's name. Then press 'Reset' from 'Tenant payment' tab. It resests your entire sheet with all saved changes.
