const fs = require('fs')
// const request = require('request')
const got = require('got')
var excel = require('exceljs')
const { domain } = require('process')

const sourceWebparts  = "output/output-webpart-other-external.txt"
const sourceMain      = "output/output-main-other-external.txt"
const sourcePagesAudit = "output/output-pageaudit-pages.txt"      // Get all pages so we can find parent page status

var validLinksFile   = "output/output-valid-links-external.txt"
var brokenLinksFile  = "output/output-broken-links-external.txt"
var validDocsFile    = "output/output-valid-external.txt"
var unusedDocsFile   = "output/output-unused-external.txt"

const outputFile    = "output/GatewayExternalLinksAnalysis.xlsx"  

var ts = new Date();
console.log(ts.toLocaleTimeString(), " :: Start  - Compare External Links")


// fs.writeFileSync(validLinksFile, "");
// fs.writeFileSync(brokenLinksFile, "");
// fs.writeFileSync(validDocsFile, "");
// fs.writeFileSync(unusedDocsFile, "");


loadData();
analyseData().then(() => {
  var ts = new Date();
  console.log(ts.toLocaleTimeString(), " :: Finish - Compare External Links")
});


function loadData() {


  // Get all Links 

  fileContents = fs.readFileSync(sourceMain, "utf8"); // Read in the URLs from the MAIN file
  linksInContent = fileContents.split("\n");
  linksInContent.forEach(stringToArray)
  linksInContent = linksInContent.filter(function (el) { return el != ""; })
  console.log("    Links in Main Content file:             " + linksInContent.length);

  fileContents = fs.readFileSync(sourceWebparts, "utf8");  // Read in URLs from the WEBPARTS file
  linksInWebParts = fileContents.split("\n");
  linksInWebParts.forEach(stringToArray)
  linksInWebParts = linksInWebParts.filter(function (el) { return el != ""; })
  console.log("    Links in Web Parts file:                 " + linksInWebParts.length);

  allLinks = linksInWebParts.concat(linksInContent) // Merge Webparts and Main into a single sorted Array
  allLinks = Array.from((new Map(allLinks.map((item) => [item.join(), item]))).values());  // Remove duplicates
  allLinks.sort();                                                                         // Sort
  console.log("    Total External Links:                   " + allLinks.length + "\n");

  // Get all Pages
  fileContents = fs.readFileSync(sourcePagesAudit, "utf8"); // Read in the Page from the PAGEAUDIT file
  allPages = fileContents.split("\n");
  allPages.forEach(stringToArray)
  allPages = allPages.filter(function (el) { return el != ""; })


  function stringToArray(sourceStr, index, arr) {
    sourceStr = sourceStr.replace(/["]+/g, '')
    resArr = sourceStr.split(" | ")
    arr[index] = resArr
  }
}

//  
//  NOW FOR THE ANALYSIS !!!
//  ========================
//  
//  To produce the following reports:
// 
//  (1) A list of all valid links that go to files held in SharePoint, along with the page they are on and the file published status
//        The links to published documents are good
//        The links to unpublished documents are broken for most users and either need to be removed, or the documents need to be published
//  (2) A list of all links whick are broken - the file they point to doesn't exist in SharePoint
//        The links need to be removed
//  (3) A list of all valid files that have 1 or more links to them, along with their published status
//        The published files with links to them are good
//        The unpublished files with links to them are broken for most users (this is covered by action (1)
//  (4) A list of all files that have no links going to them, along with their published status
//        The published files need to unpublished
//        The unpublished files are fine as they are
//


async function analyseData() {

  
    // Open the log file
    var outputLog = "output/outputLog.txt"
    fs.writeFileSync(outputLog, "");

  var workbook = new excel.Workbook();

  var summarySheet = workbook.addWorksheet('Summary');
  summarySheet.columns = [
    { header: 'Summary',   key: 'summary',   width: 110 },
    { header: 'Total',     key: 'total',     width:  10 },
    { header: 'Published', key: 'published', width:  10}
  ];

  var validLinksSheet = workbook.addWorksheet('Valid Links');
  validLinksSheet.columns = [
    { header: 'External URL',                  key: 'fileurl',                    width: 70 },
    { header: 'Parent Page URL',               key: 'pageurl',                    width: 70 },
    { header: 'Location in Parent Page',       key: 'location',                   width: 20 },
    { header: 'Link Status',                   key: 'linkstatus',                 width: 20 },    
    { header: 'Parent Page Moderation Status', key: 'parentpagemoderationstatus', width: 20 },
    { header: 'Parent Page Published Level',   key: 'parentpagepublishedlevel',   width: 20 },
    { header: 'File Type',                     key: 'filetype',                   width: 10 }
  ];


  // Summary Text
  summarySheet.addRow(["List all external links from Gateway, and identify whether they are from published pages"]).commit();
  summarySheet.addRow([" "]).commit();

  // Get all links that go to a file held in SharePoint
  var validLinks = compareArrays(allLinks, allPages, "valid links");
  var pubPages = countPublishedPages(validLinks, 6);
  await checkUrls(validLinks)
  console.log("after checkUrl")
  summarySheet.addRow(["All External Links:",validLinks.length, pubPages]).commit();
  console.log("    All External Links:                     " + validLinks.length + "   " + pubPages);
  validLinks.forEach(OutputValidLinks)



  // Create the Excel file
  await workbook.xlsx.writeFile(outputFile);

  
  // Write to Excel and to txt files
  
  // Columns in source Arrays
  //                       FILES    LINKS
  // 0  : URL                X        X
  // 1  : Parent Page                 X
  // 2  : Location                    X 
  // 3  : Modified Date       X       Use File Moderation Status copied from File array
  // 4  : Modified By         X       Use Parent Page Moderation Status copied from Page Array
  // 5  : Moderation Status   X       Use File Published Version copied from File Array
  // 6  : Publishing Level    X       Use Parent Page Pubished Version copied from Page Array
  // 7  : Checkout User       X
  // 8  : Version             X
  // 9  : Published Version   X
  // 10 : FileType            X       X


 
  function OutputValidLinks(record) {
    validLinksSheet.addRow([record[0],record[1],record[2],record[3],record[4],record[6],record[10]]).commit();
  };

  function countPublishedPages(arrayA, column){
    var pubPagesCount = 0;
    for (i = 0; i < arrayA.length; i++) {
      if ((arrayA[i][column] != "New") && (arrayA[i][column] != "Unpublished")){
        pubPagesCount++
      }
    }
    return pubPagesCount
  }

  function getDomain(record){
    parts = record.split("//")       // array containing before 
    parts2 = parts[1].split("/")
    return parts2[0]
  }


  async function checkUrls(recordArray) {
    var skipDomain = ""

    //for (i = 0; i < recordArray.length; i++) {
    //  var ts = new Date();
    //  outputStr = ts.toLocaleTimeString() + " :: " + i + " :: " + recordArray[i][0]
    //  console.log(outputStr)
    //  fs.appendFileSync(outputLog, outputStr + "\n");
    //}

    for (i = 0; i < recordArray.length; i++) {

      if (recordArray[i][0].substring(0, 4).toLowerCase() === "http") {

        var domain = getDomain(recordArray[i][0].toLowerCase())  // get each domain
        // console.log(domain)
        if (domain !== skipDomain) {                         // skip if domain same as previous timeout

          if (i == 0) {
            try {
              const response = await got(recordArray[i][0], { retry: 0,  timeout: 9000 });  //timeout 9 sec
              recordArray[i][3] = response.statusCode
            } catch (error) {
              if (typeof error.response !== 'undefined') {
                recordArray[i][3] = error.response.statusCode
              } else {
                recordArray[i][3] = 0
              }
            }
          }
          if (i > 0) {

            if (recordArray[i][0].toLowerCase() === recordArray[i - 1][0].toLowerCase()) {    // see whether URL has been checked & copy result
              recordArray[i][3] = recordArray[i - 1][3]
              logData(i, recordArray[i])
            } else {
              try {
                const response = await got(recordArray[i][0], { retry: 0, timeout: 9000 });  //timeout 9 sec
                recordArray[i][3] = response.statusCode
                skipDomain = ""
                logData(i, recordArray[i])
              } catch (error) {
                if (typeof error.response !== 'undefined') {
                  recordArray[i][3] = error.response.statusCode
                } else {
                  recordArray[i][3] = "Timeout"
                  skipDomain = domain;
                }
                logData(i, recordArray[i])
              }
            }
          }
        } else {
          recordArray[i][3] = "Skip"
          logData(i, recordArray[i])
        }
      } else {
        recordArray[i][3] = "Not http(s)"
        logData(i, recordArray[i])
      }
    }
  }

  function logData( index, record) {
    var ts = new Date();
    outputStr = ts.toLocaleTimeString() + " :: " + index + " :: " + record[3] + " :: " + record[0]
    fs.appendFileSync(outputLog, outputStr + "\n");
    if (i % 50 == 0) {
      console.log(outputStr)
    }
  }

  function oldcheckUrl(url) {
    // console.log(url)
    request(url, function (error, response, body) {
      if (error) {
        return '0'
      } else {
        console.log(response.statusCode)
        return response.statusCode;
      }
    });
  }

  async function old2checkUrl(url) {
    try {
      const response = await got(url);
      console.log(response.statusCode)
      return response.statusCode;
    } catch (error) {
      // console.log(error.response.body);
      //=> 'Internal server error ...'
      return '0'
    }
  }

  async function checkUrl(element) {
    try {
      const response = await got(element[0]);
      console.log(response.statusCode)
      return response.statusCode;
    } catch (error) {
      // console.log(error.response.body);
      //=> 'Internal server error ...'
      return '0'
    }
  }





  function compareArrays(arrayA, arrayB, option) {
    var res = []

    switch (option) {

      case "valid links":   // All external links
        for (i = 0; i < arrayA.length; i++) {
          found = false;
          for (j = 0; j < arrayB.length; j++) {
            if (arrayA[i][1].toLowerCase() === arrayB[j][0].toLowerCase()) {
              found = true
              arrayA[i][4] = arrayB[j][5]  // Cpoy across the parent page moderation status field
              arrayA[i][6] = arrayB[j][9]  // Cpoy across the parent page Published Version field 
            }
          }
          res.push(arrayA[i])
        }
        return res

    }

  }
}