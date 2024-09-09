package main

import (
	"context"
	"fmt"
	"log"
	"regexp"
	"strings"
	"time"

	"github.com/PuerkitoBio/goquery"
	"github.com/chromedp/chromedp"
	"github.com/gocolly/colly/v2"
	"github.com/xuri/excelize/v2"
)

type ProductInfo struct {
	Name                        string
	Category                    string
	Price                       string
	AvailableSizes              []string
	Images                      []string
	Breadcrumbs                 []string
	URL                         string
	SizeFit                     string
	PopupImage                  string
	PopupPrice                  string
	TitleOfDescription          string
	GeneralDescriptionOfProduct string
	ProductDetails              ProductDetails
	Tags                        []string
}

type ProductDetails struct {
	Features  []string
	ArticleID string
	Color     string
	Country   string
}

func scrapeProductDetails(productURL string, file *excelize.File, rowIndex int) {
	c := colly.NewCollector()

	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL.String())
	})

	c.OnHTML("body", func(e *colly.HTMLElement) {
		product := ProductInfo{}

		product.Name = e.ChildText("h1.itemTitle.test-itemTitle")

		product.Category = e.ChildText("a.groupName span.categoryName.test-categoryName")

		product.Price = e.ChildText("span.price-value.test-price-value")

		e.ForEach("li.sizeSelectorListItem button.sizeSelectorListItemButton", func(_ int, el *colly.HTMLElement) {
			size := el.Text

			product.AvailableSizes = append(product.AvailableSizes, size)
		})

		e.ForEach("ul.selectable-image-group li.selectableImageListItem img.selectableImage", func(_ int, el *colly.HTMLElement) {


			src := el.Attr("src")

			imageURL := "https://shop.adidas.jp" + src

			product.Images = append(product.Images, imageURL)
		})

		e.ForEach("div.breadcrumb_wrap ul.breadcrumbList li.breadcrumbListItem a.breadcrumbListItemLink", func(_ int, el *colly.HTMLElement) {
			breadcrumbText := el.Text
			product.Breadcrumbs = append(product.Breadcrumbs, breadcrumbText)
		})

		product.URL = e.Request.URL.String()

		e.ForEach("div.sizeFitBar div.label span", func(_ int, el *colly.HTMLElement) {
			sizeFitText := el.Text
			
			product.SizeFit += sizeFitText + " "
		})

		e.ForEach("div.coordinate_item_tile.test-coordinate_item_tile", func(_ int, el *colly.HTMLElement) {
			popupImage := el.ChildAttr("div.coordinate_image img.coordinate_image_body", "src")
			if popupImage != "" {
				product.PopupImage = "https://shop.adidas.jp" + popupImage
			}
			popupPrice := el.ChildText("div.coordinate_price span.price-value.test-price-value")
			if popupPrice != "" {
				product.PopupPrice = popupPrice
			}
		})

		product.TitleOfDescription = e.ChildText("h4.heading.itemFeature.test-commentItem-subheading")
		product.GeneralDescriptionOfProduct = e.ChildText("div.commentItem-mainText.test-commentItem-mainText")

		productDetails := ProductDetails{}

		e.ForEach("ul.articleFeatures.description_part.css-1lxspbu li.articleFeaturesItem", func(_ int, el *colly.HTMLElement) {
			liText := el.Text

			switch {
			case el.DOM.HasClass("test-feature"):
				productDetails.Features = append(productDetails.Features, liText)
			case el.DOM.HasClass("test-articleId"):
				productDetails.ArticleID = el.ChildText(".test-itemComment-article")
			case el.DOM.HasClass("test-itemColor"):
				productDetails.Color = liText
			case el.DOM.HasClass("test-itemCountry"):
				productDetails.Country = liText
			}
		})

		product.ProductDetails = productDetails

		e.ForEach("div.test-category_link a", func(_ int, el *colly.HTMLElement) {
			tag := el.Text
			product.Tags = append(product.Tags, tag)
		})

		rowData := []interface{}{
			product.Name,
			product.Category,
			product.Price,
			strings.Join(product.AvailableSizes, ", "),
			strings.Join(product.Images, "; "),
			strings.Join(product.Breadcrumbs, " > "),
			product.URL,
			product.SizeFit,
			product.PopupImage,
			product.PopupPrice,
			product.TitleOfDescription,
			product.GeneralDescriptionOfProduct,
			strings.Join(product.ProductDetails.Features, ", "),
			product.ProductDetails.ArticleID,
			product.ProductDetails.Color,
			product.ProductDetails.Country,
			strings.Join(product.Tags, ", "),
			"",
			"",
			"",
			"",
			"",
			"",
		}

		sheet := "Sheet1"
		for colIndex, value := range rowData {
			cell, _ := excelize.ColumnNumberToName(colIndex + 1)
			file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex), value)
		}

		err := file.SetColWidth(sheet, "A", "Q", 100)
		if err != nil {
			log.Fatalf("Error setting column width: %v", err)
		}

		err = file.SetColWidth(sheet, "Q", "S", 250)
		if err != nil {
			log.Fatalf("Error setting column width: %v", err)
		}

		err = file.SetColWidth(sheet, "T", "T", 250)
		if err != nil {
			log.Fatalf("Error setting column width: %v", err)
		}

		err = file.SetColWidth(sheet, "U", "W", 20)
		if err != nil {
			log.Fatalf("Error setting column width: %v", err)
		}

		err = file.SetColWidth(sheet, "R", "S", 40)
		if err != nil {
			log.Fatalf("Error setting column width: %v", err)
		}

		err = file.SetRowHeight(sheet, 1, 30)
		if err != nil {
			log.Fatalf("Error setting row height: %v", err)
		}

	})

	c.OnError(func(_ *colly.Response, err error) {
		log.Println("Something went wrong:", err)
	})

	err := c.Visit(productURL)
	if err != nil {
		log.Println("Failed to visit product URL:", err)
	}

}

func scrapeCategoryPages(categoryLink string, file *excelize.File, mainFile string) {
	c := colly.NewCollector(
		colly.AllowedDomains("shop.adidas.jp"),
	)

	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting", r.URL.String())
	})

	var productLinks []string
	productLinkRegex := regexp.MustCompile(`^https://shop\.adidas\.jp/products/[^/]+$`)

	c.OnHTML("a[href]", func(e *colly.HTMLElement) {
		link := e.Request.AbsoluteURL(e.Attr("href"))

		if productLinkRegex.MatchString(link) {
			productLinks = append(productLinks, link)
		}
	})

	c.OnHTML("body", func(e *colly.HTMLElement) {
		if len(productLinks) == 0 {
			fmt.Println("No products found on this page, skipping...")
			return
		}

		for i, link := range productLinks {
			if i >= 300 {
				break
			}

			fmt.Printf("Processing product link %d: %s\n", i+1, link)

			rowIndex := i + 2
			scrapeProductDetails(link, file, rowIndex)
			err := file.SaveAs(mainFile)
			if err != nil {
				log.Fatal("Could not save file:", err)
			} else {
				fmt.Printf("File saved success!! %d: %s\n", i+1, link)

			}

		}
	})

	c.OnError(func(_ *colly.Response, err error) {
		log.Println("Something went wrong:", err)
	})

	for page := 1; page <= 30; page++ {
		pageURL := categoryLink
		if page > 1 {
			pageURL = fmt.Sprintf("%s&page=%d", categoryLink, page)
		}
		if !strings.HasPrefix(pageURL, "https://") {
			pageURL = "https://shop.adidas.jp" + pageURL
		}
		fmt.Println("Visiting category page:", pageURL)
		err := c.Visit(pageURL)
		if err != nil {
			log.Println("Failed to visit page:", err)
		}
	}
}

func writeExcelHeader(file *excelize.File) {
	sheet := "Sheet1"
	headers := []string{"Name", "Category", "Price", "Available Sizes", "Images", "Breadcrumbs", "URL", "SizeFit", "PopupImage", "PopupPrice", "TitleOfDescription", "GeneralDescription", "Features", "ArticleID", "Color", "Country", "Tags", "SizeChartHeader", "FirstRows", "AllRows", "Recommendation", "ReviewsCount", "AvgRating"}
	for colIndex, header := range headers {
		cell, _ := excelize.ColumnNumberToName(colIndex + 1)
		file.SetCellValue(sheet, fmt.Sprintf("%s1", cell), header)
	}
}

func readExcel(filePath string) []string {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Error opening file: %v", err)
	}
	defer f.Close()

	sheet := "Sheet1"

	rows, err := f.GetRows(sheet)
	if err != nil {
		log.Fatalf("Error getting rows: %v", err)
	}

	productUrls := make([]string, 0)
	for rowIndex, row := range rows {
		if rowIndex == 0 {
			continue
		}
		if len(row) < 7 {
			continue
		}
		productURL := row[6]
		productUrls = append(productUrls, productURL)
	}
	return productUrls
}

func scrapeDynamicContent(url string) ([]string, []string, [][]string, string, string, string) {
	ctx, cancel := context.WithTimeout(context.Background(), 4*time.Minute)
	defer cancel()

	opts := append(chromedp.DefaultExecAllocatorOptions[:],
		chromedp.ExecPath("/usr/bin/google-chrome"),
		chromedp.Flag("headless", false),
		chromedp.Flag("disable-gpu", true),
	)

	allocCtx, cancelAlloc := chromedp.NewExecAllocator(ctx, opts...)
	defer cancelAlloc()

	chromeCtx, cancelChrome := chromedp.NewContext(allocCtx)
	defer cancelChrome()

	var sizeChartContent, avgScore, reviewsCount, recommendation string

	err := chromedp.Run(chromeCtx,
		chromedp.Navigate(url),
		chromedp.WaitVisible(`body`, chromedp.ByQuery),
		chromedp.Sleep(10*time.Second),

		chromedp.Click("body", chromedp.ByQuery),
		chromedp.Sleep(1*time.Second),

		chromedp.Evaluate(`document.querySelector("div#modalArea.modalArea")?.remove();`, nil),
		chromedp.Sleep(2*time.Second),

		chromedp.ActionFunc(func(ctx context.Context) error {
			return performSmoothScroll(ctx)
		}),

		chromedp.WaitVisible(`h3.sizeDescriptionWrapHeading`, chromedp.ByQuery),
		chromedp.ScrollIntoView(`h3.sizeDescriptionWrapHeading`, chromedp.ByQuery),
		chromedp.OuterHTML(`div.sizeChart`, &sizeChartContent),
	)

	if err != nil {
		fmt.Printf("Error scraping product at URL: %s. Error: %v\n", url, err)
		return nil, nil, nil, "", "", ""
	}

	if err := chromedp.Run(chromeCtx, chromedp.Text(`div.BVRRRatingNormalOutOf span.BVRRNumber.BVRRRatingNumber`, &avgScore, chromedp.ByQuery)); err != nil {
		fmt.Println("avgScore count element not found, assigning default value.")
		avgScore = "N/A"
	}

	if err := chromedp.Run(chromeCtx, chromedp.Text(`span.BVRRNumber.BVRRBuyAgainTotal`, &reviewsCount, chromedp.ByQuery)); err != nil {
		fmt.Println("Reviews count element not found, assigning default value.")
		reviewsCount = "N/A"
	}

	if err := chromedp.Run(chromeCtx, chromedp.Text(`span.BVRRBuyAgainPercentage span.BVRRNumber`, &avgScore, chromedp.ByQuery)); err != nil {
		fmt.Println("recommendation count element not found, assigning default value.")
		recommendation = "N/A"
	}

	doc, err := goquery.NewDocumentFromReader(strings.NewReader(sizeChartContent))
	if err != nil {
		fmt.Printf("Error loading size chart for URL: %s. Error: %v\n", url, err)
		return nil, nil, nil, "", "", ""
	}

	headers := []string{}
	doc.Find("thead.sizeChartTHeader th").Each(func(_ int, th *goquery.Selection) {
		headerText := th.Text()
		if headerText != "" {
			headers = append(headers, headerText)
		}
	})

	var allRows [][]string

	doc.Find("tbody tr").Each(func(rowIndex int, tr *goquery.Selection) {
		rowData := []string{}
		tr.Find("td").Each(func(_ int, td *goquery.Selection) {
			rowData = append(rowData, td.Text())
		})
		allRows = append(allRows, rowData)
	})

	if len(headers) == 0 || len(allRows) == 0 {
		fmt.Printf("No size chart data found for URL: %s. Skipping...\n", url)
		return nil, nil, nil, avgScore, reviewsCount, recommendation
	}

	return allRows[0], headers, allRows, avgScore, reviewsCount, recommendation
}

func performSmoothScroll(ctx context.Context) error {
	for i := 0; i < 10; i++ {
		err := chromedp.Run(ctx,
			chromedp.Evaluate(`window.scrollBy(0, window.innerHeight / 2);`, nil),
			chromedp.Sleep(500*time.Millisecond),
		)
		if err != nil {
			return err
		}
	}
	return nil
}

func scrapeSizeChartsAndUpdateExcel(urls []string, filePath string) {
	file, err := excelize.OpenFile(filePath)
	if err != nil {
		log.Fatalf("Error opening Excel file: %v", err)
	}
	defer file.Close()

	sheet := "Sheet1"

	sizeChartData := make(map[string]struct {
		Header         []string
		FirstRow       []string
		AllRows        [][]string
		AvgScore       string
		ReviewsCount   string
		Recommendation string
	})

	for _, url := range urls {
		headers, rowsData, allRows, avgScore, reviewsCount, recommendation := scrapeDynamicContent(url)
		if headers == nil {
			continue
		}



		sizeChartData[url] = struct {
			Header         []string
			FirstRow       []string
			AllRows        [][]string
			AvgScore       string
			ReviewsCount   string
			Recommendation string
		}{
			Header:         headers,
			FirstRow:       rowsData,
			AllRows:        allRows,
			AvgScore:       avgScore,
			ReviewsCount:   reviewsCount,
			Recommendation: recommendation,
		}

		err = updateExcelWithSizeChart(file, sheet, filePath, sizeChartData)
		if err != nil {
			log.Fatalf("Error updating spreadsheet: %v", err)
		}

		err = file.SaveAs(filePath)
		if err != nil {
			log.Fatalf("Error saving updated Excel file: %v", err)
		}
	}

}

func formatAllRows(allRows [][]string) string {
	var formattedRows []string
	for _, row := range allRows {
		formattedRow := strings.Join(row, ",")
		formattedRows = append(formattedRows, formattedRow)
	}
	return strings.Join(formattedRows, ";")
}

func updateExcelWithSizeChart(file *excelize.File, sheet, filePath string, sizeChartData map[string]struct {
	Header         []string
	FirstRow       []string
	AllRows        [][]string
	AvgScore       string
	ReviewsCount   string
	Recommendation string
}) error {
	rows, err := file.GetRows(sheet)
	if err != nil {
		return fmt.Errorf("error getting rows: %v", err)
	}


	for url, data := range sizeChartData {
		rowFound := false
		for rowIndex, row := range rows {
			if rowIndex == 0 {
				continue
			}
			if len(row) <= 6 {
				continue
			}
			if row[6] == url {

				cell, _ := excelize.ColumnNumberToName(18)
				file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), strings.Join(data.Header, ", "))

				cell, _ = excelize.ColumnNumberToName(19)

				if len(data.FirstRow) > 0 {
					file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), strings.Join(data.FirstRow, ", "))
				}

				cell, _ = excelize.ColumnNumberToName(20)
				allRowsFormatted := formatAllRows(data.AllRows)

				file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), allRowsFormatted)

				cell, _ = excelize.ColumnNumberToName(21)
				file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), data.AvgScore)

				cell, _ = excelize.ColumnNumberToName(22)
				file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), data.ReviewsCount)

				cell, _ = excelize.ColumnNumberToName(23)
				file.SetCellValue(sheet, fmt.Sprintf("%s%d", cell, rowIndex+1), data.Recommendation)

				rowFound = true
				break
			}
		}

		if !rowFound {
			log.Printf("URL not found in sheet: %s", url)
		}


		err = file.SaveAs(filePath)
		if err != nil {
			return fmt.Errorf("error saving file: %v", err)
		}
	}
	return nil
}

func main() {
	FILE := "products.xlsx"

	file := excelize.NewFile()
	defer file.Close()

	writeExcelHeader(file)

	categoryLink := "https://shop.adidas.jp/item/?category=wear&condition=6&limit=120&page="
	scrapeCategoryPages(categoryLink, file, FILE)

	productsUrls := readExcel(FILE)

	scrapeSizeChartsAndUpdateExcel(productsUrls, FILE)

}
