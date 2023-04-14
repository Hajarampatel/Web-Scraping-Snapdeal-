const puppeteer = require("puppeteer");
const stringSimilarity = require("string-similarity");
const reader = require('xlsx')
const Output = [];    // storing data of books with low price 

(async () => {
  let data = []        // storing data from Input.xlsx
  const file = reader.readFile("./Input.xlsx")
  const sheets = file.SheetNames
  for(let i = 0; i < sheets.length; i++){
      const temp = reader.utils.sheet_to_json(
      file.Sheets[file.SheetNames[i]])
      temp.forEach((res) => {
      data.push(res)
   })
  }

 const browser  = await puppeteer.launch({headless : false, executablePath : "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"});
 const page = await  browser.newPage();
 await page.setDefaultNavigationTimeout(0);
          let books_isbn = []           // ISBN from Input.xlsx
          let books_title = []          // Title from Input.xlsx
         data.forEach( async (arrayItem) => { 
          books_title.push(arrayItem.Book_Title)
          books_isbn.push(arrayItem.ISBN)
        });
        
  for (let index = 0; index < books_isbn.length; index++) { // Iterating for each data in Input.xlsx
        
      let lowPriceBook = {       // Initializing  object for store data with low price
        ISBN : "NO",
        title : "NO",
        price : "NO",
        author : "NO",
        pageUrl : "NO",
        publisher : "NO",
        found :  "NO"
      }
    await page.goto("https://www.snapdeal.com/", {timeout: 0}); 
    await page.type("#inputValEnter", JSON.stringify(books_isbn[index]));
    await page.click("#sdHeader > div.headerBar.reset-padding > div.topBar.top-bar-homepage.top-freeze-reference-point > div > div.col-xs-14.search-box-wrapper > button > span");
    await page.waitForNavigation();

     const excelTitle = books_title[index].toLowerCase()
     lowPriceBook.title = books_title[index]
     lowPriceBook.ISBN = books_isbn[index]


      const books = await page.$$(".product-desc-rating")
      if(books.length !== 0){ 
        lowPriceBook.found = "YES"
     }
     
      let details = []           //store data of each book found on snapdeal for perticular ISBN so we can compare price for that book
        
       for(const  p of books){   // Initializing  object for store data of books in details[]
        let book = {
            title : null,
            price : null,
            author : null,
            pageUrl : null
        }
                              // p = Iterating for each  book on snapdeal 
        let match = null
        try {
           book.title = await page.evaluate((el) => el.querySelector("a > .product-title").title, p)
           match = stringSimilarity.compareTwoStrings(excelTitle.replace(/ \([\s\S]*?\)/g, '').toLowerCase(), book.title.replace(/ \([\s\S]*?\)/g, '').toLowerCase())
           if(match < .9) continue       // matching input title with snapdeal title with >= 90%
        } catch (error) {  }
        try {
          const tempPrice  = await page.evaluate((element) => element.querySelector(".lfloat .product-price").innerHTML.slice(5), p)
          book.price = parseInt(tempPrice)
        } catch (error) {  }
        try {
             book.author =  await page.evaluate((element) => element.querySelector(".product-author-name").title, p) 
        } catch (error) {book.author = "NA" }
        try {
          book.pageUrl =  await page.evaluate((element) => element.querySelector("a").href, p) 
        } catch (error) { }
        details.push(book)
      }

        let sortedBooks = details.sort(
          (p1, p2) => (p1.price > p2.price) ? 1 : (p1.price < p2.price) ? -1 : 0); // shorting books by price in increasing order
          
          lowPriceBook.price = sortedBooks[0].price        
          lowPriceBook.author = sortedBooks[0].author
          lowPriceBook.pageUrl = sortedBooks[0].pageUrl

         await page.goto(sortedBooks[0].pageUrl)

         try {
           const openbook = await page.$$("#productOverview > div.col-xs-14.right-card-zoom.reset-padding > div > div.pdp-elec-topcenter-inner.layout > div.highlightsTileContent.highlightsTileContentTop.clearfix")
            for (const item of openbook) {
              var tempPublisher =  await page.evaluate((el) => el.querySelector("div > ul > li:nth-child(3) > span.h-content").innerHTML, item)
              if(JSON.stringify(tempPublisher).substring(1, 10).toLowerCase() === "publisher") {lowPriceBook.publisher = tempPublisher.slice(10)}
              else 
              {lowPriceBook.publisher = await page.evaluate((el) => el.querySelector("div > ul > li:nth-child(4) > span.h-content").innerHTML.slice(10), item)  }
            }
         } catch (error) { }

         Output.push(lowPriceBook)   // storing final low price book with title match >= 90% in final output
  }
        await browser.close();

         var workbook = reader.utils.book_new()
         var worksheet = reader.utils.json_to_sheet(Output)   // storing data in file Output.xlsx
         reader.utils.book_append_sheet(workbook, worksheet)
         reader.writeFile(workbook, "Ouptut.xlsx")

})();
