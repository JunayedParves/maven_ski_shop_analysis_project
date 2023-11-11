# Maven Ski Shop Black Friday Sales Analysis Project
I have completed this project as a part of Maven course "Python Foundation for Data Analysis". This helps me to practice my knowledge I have learned from the course. I have fun to do the projects. 

### ProjectOverview:
Last three months ski shop sales data is in excel, and
I need to analyze Black Friday sales data!
In the Excel workbook there is
missing data for taxes and totals. I have to fill in those
rows of data using Python.
I need to calculate some key metrics by aggregating
data, which will be super helpful for determining
how well the shop performed during Black Friday.

### Data Source
Sales Data: The primary dataset used for this analysis is the "maven_ski_shop_data.xlsx" file, containing detailed information about each sale made by the shop. 

### Key Objectives:
- Read in data from an Excel workbook
- Define a function that prints cell contents
- Create a dictionary using a comprehension and string methods
- Use for loops to manipulate Excel data
- Import and call a previously saved function
- Write data into Excel cells
- Save an Excel workbook
- Define a function that sums Excel columns, leveraging a list comprehension
- Apply numerical functions to calculate KPIs
- Use set operations to find unique items
- Create a dictionary using nested loops
- Challenge: Write a function that calculates the sum of an Excel column, grouped by the unique values in another column


### Tools
- Python (jupyter notebook). 
- Microsoft Excel. 

### Import Libraries & Data Preparation
In the initial stage I performed the following task:
  1. Import required libraries.
```
import openpyxl as xl
from tax_calculator import tax_calculator
# pprint prints dictionaries a bit more nicely than print
from pprint import pprint
```
  2. Load my sheet in python.
```
wb=xl.load_workbook(filename='maven_ski_shop_data.xlsx')
orders = wb['Orders_Info']
```
  3. Define column_printer function.
```
def column_printer(sheet, column):
    for i in range (1, sheet.max_row +1):
        print(f'{column}{i}', sheet[f'{column}{i}'].value)
```
  4. Create order data dictionary.
```
order_dict = {
    orders[f'A{order}'].value:[
        orders[f'B{order}'].value,
        orders[f'C{order}'].value,
        orders[f'D{order}'].value,
        orders[f'G{order}'].value,
        str(orders[f'H{order}'].value).split(', ')
    ]
    for order in range(2, orders.max_row + 1)
}
```

  5. Calculate sales tax.
```
for order in order_dict.values():
    if order[3] == 'Sun Vally':
        transaction = tax_calculator(order[2], .08)
    elif order[3] == 'Mammoth':
        transaction = tax_calculator(order[2], .0775)
    else:
        transaction = tax_calculator(order[2], .06)
    order.insert(3, transaction[1])
    order.insert(4, transaction[2])
```
  6. Write sales tax and total into the excel sheet and save it.
```
for index, order in enumerate (order_dict.values(), start =2):
    orders[f'E{index}'] = order[3]
    orders[f'F{index}'] = order[4]
#write sales tax and total into workbook
wb.save('maven_ski_shop_data_junayed.xlsx')
``` 

### Data Analysis:
  1. Sum The Subtotal, Tax, and Total Columns.
```
def column_sum(column_index, dictionary):
    return round(sum([value[column_index] for value in dictionary.values()]),2)
print(column_sum(2, order_dict))
print(column_sum(3, order_dict))
print(column_sum(4, order_dict))
```
  2. Find average of our subtotals.
```
round(column_sum(2,order_dict)/ len(order_dict),2)
```
  3. Find how many unique customers did we have.
```
unique_customers = len(set([order[0] for order in order_dict.values()]))

order_per_customer = round(len(order_dict) / unique_customers,3)

order_per_customer
```
  4. How many items in total did we sell.
```
sum([len(order[6]) for order in order_dict.values()])
```
  5. Which location perform good or bad.
```
location_sums = {}

for data in order_dict.values():
    
    location = data[5]
    
    if location not in location_sums:
        
        location_sums[location] = 0
    
    location_sums[location] += data[2]
    
location_sums
```
  6. Create a aggregator function to find date and customar-wise total sale.
```
def aggregator(category_index, field_to_sum_index, dictionary):
    
    category_sums = {}
    
    for data in dictionary.values():
        
        category = data[category_index]
        
        category_sums[category] = round(category_sums.get(category, 0) + data[field_to_sum_index],2)
        
    return category_sums
```
### Findings
  1. Average value of orders was $323.39
  2. There is 1.4 unique customer.
  3. Shop had sold 54 items during the black friday sale.
  4. Total sale by location was Sun Valley: 1268.84,
     Stowe: 3582.81,
     Mammoth: 3879.80.
     We can see Mammoth performed better and Sun vally did the worst.

### Recommendations
Based on the analysis I can recommend that we should do further analysis on Sun Vally to find 
the under performance reason to do better in futer. 




