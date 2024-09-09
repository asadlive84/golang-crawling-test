# Golang Web Scraper

This Go application is designed to scrape product details and size charts from the any static or dynamic website. It extracts various product details, including name, category, price, available sizes, images, and more, and saves them into an Excel file.

## Prerequisites

Before you can run this application, ensure that you have the following installed:

- [Go](https://golang.org/dl/) (version 1.20 or later)
- [Google Chrome](https://www.google.com/chrome/) (required for headless browsing)
- [Colly](https://github.com/gocolly/colly) (web scraping framework)
- [chromedp](https://github.com/chromedp/chromedp) (headless Chrome/Chromium driver)
- [Excelize](https://github.com/xuri/excelize) (library to handle Excel files)

## Installation

1. Clone the repository or download the source code.
2. Navigate to the project directory.
3. Install the necessary dependencies:

    ```sh
    go get -u github.com/PuerkitoBio/goquery
    go get -u github.com/chromedp/chromedp
    go get -u github.com/gocolly/colly/v2
    go get -u github.com/xuri/excelize/v2
    ```

## Usage

### Running the Application

1. Ensure that the Excel file you're working with is in the correct format with headers.
2. Open a terminal and navigate to the project directory.
3. Run the application with the following command:

    ```sh
    go mod tidy
    ```

    ```sh
    go run .
    ```

### Parameters

The application requires a few parameters to be set directly in the code:

- `categoryLink`: The URL of the Adidas category page you want to scrape.
- `mainFile`: The path to the Excel file where you want to store the scraped data.

You can update these variables directly in the `main.go` file.

### Output

The application will scrape the specified category and product pages and store the data in the specified Excel file. It will also update existing rows in the Excel file if the product URL already exists.

## Notes

- The application scrapes up to 300 products from each category page.
- The size charts and additional product information are dynamically fetched using `chromedp` for content loaded via JavaScript.
- This project includes a web scraping component that targets the Adidas website. The scraping is done purely for educational purposes to learn and practice web scraping techniques with Go. Please use the code responsibly and ensure that it complies with any applicable legal and ethical guidelines.

## Troubleshooting

- Ensure that Google Chrome is installed and accessible by `chromedp`.
- If you encounter any issues with scraping, check the URLs provided and ensure they are valid and accessible.

