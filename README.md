# Bookzone Price Tracker

A Python web scraping project that monitors book prices on [Bookzone](https://bookzone.ro/), records daily price history, and tracks the lowest recorded prices over time.

## Features

- **Automated price tracking** — extracts book titles and current prices from Bookzone.
- **Daily price history** — stores recorded prices in an Excel spreadsheet.
- **Lowest price tracking** — calculates the lowest recorded price for each book.
- **Current price overview** — displays the latest recorded price alongside the historical minimum.
- **Configurable book list** — stores book URLs in a separate JSON file, making it easy to add or remove books without modifying the Python code.
- **Excel report generation** — organizes data into two worksheets: `Prices` and `Lowest Prices`.

## Technologies Used

- Python
- Requests
- Beautiful Soup 4
- Pandas
- OpenPyXL
- JSON

## Configuration

Add or remove book URLs in `books.json`. The script loads the list automatically, so no changes to `main.py` are needed when updating the books being tracked.

## Usage

Run the script from the project directory:

```bash
python main.py
```

The script retrieves the current prices and saves the results to `Bookzone_Books.xlsx`.

The generated Excel file contains two worksheets:

- **Prices** — daily price history, with one column per book.
- **Lowest Prices** — the lowest recorded price and the latest recorded price for each book.

The script checks whether prices have already been recorded on the current day to avoid duplicate daily entries.

## Future Improvements

- Automated price-drop notifications.
- Improved error handling for unavailable products or changes to the website structure.
- More flexible scheduling and monitoring options.

## Disclaimer

This project is intended for educational purposes and personal use. It uses web scraping to collect publicly displayed product prices. Website structure and availability may change over time.
