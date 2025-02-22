# Interactive-Dashboard-Microsoft-Excel
This project was created as an interactive dashboard for a coffee shop to enable the management have an indepth undersand of their sales over time, enabling them to know sales performance of each of their coffe types, make comparison to the effect of customers having a loyalty card to their sales pattern, the size of coffer her customers demand the most and finally see sales across countries they have presence and the names of the top customers in each of these categories. 
![image](https://github.com/ChimaJerry/Interactive-Dashboard-Microsoft-Excel/assets/132655711/5723395f-6c6a-4eb0-92c7-08deb12c816b)


## STEPS TAKEN
1. ###### XLOOKUP FUNCTION AND INDEX FONCTION
    Using the customer ID as the primary key, I made use of the XLOOKUP and INDEX function of excel to find and fill in other columns that was not provided within a single sheet like the customer name, email, coffee type, Roast type and size.
   
   For the cusomer name the  following Xlookup was used:
   ```=XLOOKUP([@[Customer ID]],Customers!$A$1:$A$1001,Customers!$B$1:$B$1001,,0);```
   
   For the emails this was used =IF(XLOOKUP(C3,Customers!$A$1:$A$1001,Customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C3,Customers!$A$1:$A$1001,Customers!$C$1:$C$1001,,0))

   For the Coffee type this was used =INDEX(Products!$A$1:$G$49,MATCH(Order!$D3,Products!$A$1:$A$49,0),MATCH(Order!I$1,Products!$A$1:$G$1,0))
3. ###### MULTIPLICATION FORMULAE FOR SALES.
   The information for the sales was not directly provide from the original dataset. I derived this by multiply the unit price with the quantity of item sold.
4. ###### MULTIPLE IF FUNCTIONS
   The coffee name of the coffee type and roast type was abbreviated I created a new column, used the multiple if functions to rename the abbreviations to captured the new desired name I wished for and drag down to fill the column. 
5. ###### DATE FORMATING
    Use the cell formatting function to customise the data to show day/month/year.
6. ###### NUMBER FORMATING
   Use the custom function to add the inscription 'kg' on the size column to make it easy for readers to understand. And also inserted the $ sign into all the values in the Unit price and sales  column for a better comprehension.
7. ###### CHECKED FOR DUPLICATE VALUES
   When the data sections of MS excel, highlighted  every data provided in each cell and checked using the remove duplicated function. No duplicate were found.
8. ###### CONVERTED THE RANGE TO TABLE
   Pressed ctrl and T to create a table and bring in the table design session into the table. This is necessary because I want to use a pivot table and when the table designed is been used, if there is a new column inserted or an update to the formulae it will be so 
   much easier for the pivot table to update automatically.
9. ###### PIVOT TABLE AND PIVOT CHARTS
   Inserted a pivot table to be able to calculate for the totoal sales across each months from January 2019-December 2022 and inserted a line chart to visualized this as well, using the pivot table created another sheet to calculate the the country with the highes sales 
   and the top 5 customers and arrange them in a descending order, for the  sales by country I customixed it with the highes country showing dark green and the least country in terms of sales was made to be light green.
10. ###### INSERTED TIMELINE, SLICERS AND FORMATTED
   On the Total sales sheet went to the pivot table analyzer and inserted timeline when each order was made to enable navigation on time series when each order was made. Went ahead to insert slicers fro layalty cards, Roast type name and size ensuring they are 
   customized appropriately
11. ###### BUILDING THE DASHBOARD
    Created another sheet solely for the purpose to bring all the developed data charts, timeline and slicers together to begin visualizing the dashboard. On the Dashboard sheet inserted a rectangular shape customized it to purple colour and font white to eb able write 
    the desired name of the dashboard. Cut and pasted each of the chart and slicers into the dashboard sheet and connected the sales by country and top 5 customers chart to the timeline and slicers to have this as a one functional document.








