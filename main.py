# Import libraries
from bs4 import BeautifulSoup
from datetime import datetime
import json
import pandas as pd
from pathlib import Path
import requests

# Load the book URLs from the JSON configuration file
with open("books.json", "r", encoding="utf-8") as file:
    data = json.load(file)

books_url = data["books"]

# Check whether prices have already been recorded today
def already_sent_today():
    file_path = Path("Bookzone_Books.xlsx")

    # If the Excel file does not exist, prices have not been recorded yet
    if not file_path.exists():
        return False

    # Read the price history from the Prices worksheet
    df = pd.read_excel("Bookzone_Books.xlsx")

    # Convert the Date column to datetime objects
    df["Date"] = pd.to_datetime(df["Date"])

    # Compare the most recent recorded date with today's date
    last_date = df["Date"].max().date()
    today_date = datetime.today().date()

    return last_date == today_date

# Set the HTTP headers used when requesting the book pages
headers = {"ACCEPT": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "USER-AGENT":"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:143.0) Gecko/20100101 Firefox/143.0",}

# Store each book's title and price
books_details = []
# Scrape prices only if today's prices have not been recorded yet
if not already_sent_today():

    for url in books_url:
        # Send an HTTP request to the book page
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        web_page = response.text

        # Parse the HTML page
        soup = BeautifulSoup(web_page, "html.parser")

        # Extract the book title
        book_name = soup.find(class_="details__title__name").get_text(strip=True)

        # Extract the current book price
        book_price = soup.find(class_="details__info__price__new").get_text(strip=True)

        # Remove the currency label and convert the price to a number
        book_price = book_price.replace(" Lei", "")
        book_price = float(book_price)

        # Add the title and price to the list
        books_details.append([book_name, book_price])

    # Get today's date
    today_date = datetime.today().strftime("%Y-%m-%d")

    # Separate the book titles and prices into two lists
    book_names = [book[0] for book in books_details]
    book_prices = [book[1] for book in books_details]

    # Create a DataFrame containing today's prices
    df = pd.DataFrame(
        [book_prices],
        index=[today_date],
        columns=book_names
    )

    # Name the index column Date
    df.index.name = "Date"

    # Define the Excel file path
    file_path = Path("Bookzone_Books.xlsx")

    # If the Excel file already exists, load its price history
    if file_path.exists():
        old_df = pd.read_excel(
            file_path,
            sheet_name="Prices",
            index_col=0
        )

        # Convert the indexes to datetime objects
        old_df.index = pd.to_datetime(old_df.index)
        df.index = pd.to_datetime(df.index)

        # Combine the old price history with today's prices
        # New books are added as columns automatically
        df = pd.concat([old_df, df], sort=False)

    # Calculate the lowest recorded price for each book
    lowest_prices = df.min()

    # Convert the Series into a DataFrame
    lowest_prices_df = lowest_prices.reset_index()
    lowest_prices_df.columns = ["Book", "Lowest Price"]

    # Get the prices from the most recent row
    today_prices = df.iloc[-1]

    # Add today's prices as a new column
    lowest_prices_df["Today Price"] = today_prices.values

    # Save the price history and summary to the Excel file
    with pd.ExcelWriter("Bookzone_Books.xlsx") as writer:

        # Save all recorded prices
        df.to_excel(writer, sheet_name="Prices")

        # Save the lowest and current prices
        lowest_prices_df.to_excel(
            writer,
            sheet_name="Lowest Prices",
            index=False
        )

        # Automatically adjust column widths in both worksheets
        for worksheet in writer.book.worksheets:
            for column in worksheet.columns:
                max_length = 0

                for cell in column:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))

                worksheet.column_dimensions[column[0].column_letter].width = max_length + 2