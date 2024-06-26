# Amazon Product Searcher

This Python program uses the tkinter library to create a graphical user interface (GUI) that allows users to search for products on Amazon, display the search results in the GUI with information about the last 10 items from the first search page, save the product information in a SQL Server database, and generate a report in Excel (.xls) format of the searched products.

## Prerequisites

- Python 3.x
- Libraries: tkinter, PIL, requests, xlwt, pyodbc, selenium

## Dependency Installation

```bash
pip install pillow requests xlwt pyodbc selenium
```

## Configuration

1. Download the appropriate `chromedriver.exe` file for your operating system and version of Chrome from https://sites.google.com/a/chromium.org/chromedriver/downloads. This file is already included for Chrome version 126.0.6478.62.
2. Place `chromedriver.exe` in the same folder as the main script.
3. Ensure you have a local SQL Server database and modify the connection string in the `save_to_db` function with your server details.

---

## Database Configuration in SQL Server

### Create the Database

1. Open SQL Server Management Studio (SSMS) and connect to your SQL server.
2. In "Object Explorer," right-click on "Databases" and select "New Database...".
3. In the "New Database" dialog box, enter the database name as `ProductsDB` and click "OK".

### Create the Table

1. In "Object Explorer," expand the `ProductsDB` database.
2. Right-click on "Tables" and select "New Table...".
3. Define the columns of the table as follows:

   - Name: Type `NVARCHAR(255)`
   - Price: Type `DECIMAL(10, 2)`
   - ImageURL: Type `NVARCHAR(255)`

4. Click on the disk icon to save the table.
5. Name the table `Products` when prompted.

---

## Usage

1. Run the .exe file, and a login window will open.
2. Enter your credentials (any of the two in the `users.txt` file) and click "Log In".
3. A window will appear to search for an item on Amazon.
4. Enter the item name and click "Search".
5. The search results will be displayed in the main window, where you can save the products to the database or generate an Excel report.

## Additional Features

- The application displays an image of Amazon on the login screen.
- Product prices are correctly formatted in the user interface and in the generated report.

## Notes

- The script uses Selenium to interact with Amazon's webpage, so adjustments may be required if Amazon changes its HTML structure.
