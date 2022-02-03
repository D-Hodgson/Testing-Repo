const fs = require('fs');
var excel = require('exceljs')

const sourceWebparts   = "output/output-webpart-other-gateway.txt"
const sourceMain       = "output/output-main-other-gateway.txt"
const sourceSitesAudit = "output/output-siteaudit-other-gateway.txt"
const sourcePagesAudit = "output/output-pageaudit-pages.txt"      // Get all pages so we can find parent page status

var validLinksFile   = "output/output-valid-links-pages.txt"
var brokenLinksFile  = "output/output-broken-links-pages.txt"
var validDocsFile    = "output/output-valid-pages.txt"
var unusedDocsFile   = "output/output-unused-pages.txt"

const outputFile    = "output/GatewaySiteAnalysis.xlsx"  

var ts = new Date();
console.log(ts.toLocaleTimeString(), " :: Start  - Compare Sites")


// fs.writeFileSync(validLinksFile, "");
// fs.writeFileSync(brokenLinksFile, "");
// fs.writeFileSync(validDocsFile, "");
// fs.writeFileSync(unusedDocsFile, "");


loadData();
analyseData();

var ts = new Date();
console.log(ts.toLocaleTimeString(), " :: Finish - Compare Sites")

function loadData() {

  // Get all Links 

  fileContents = fs.readFileSync(sourceMain, "utf8"); // Read in the URLs from the MAIN file
  linksInContent = fileContents.split("\n");
  linksInContent.forEach(stringToArray)
  linksInContent = linksInContent.filter(function (el) { return el != ""; })
  console.log("    Links in Main Content file:                     " + linksInContent.length);

  fileContents = fs.readFileSync(sourceWebparts, "utf8");  // Read in URLs from the WEBPARTS file
  linksInWebParts = fileContents.split("\n");
  linksInWebParts.forEach(stringToArray)
  linksInWebParts = linksInWebParts.filter(function (el) { return el != ""; })
  console.log("    Links in Web Parts file:                        " + linksInWebParts.length);

  allLinks = linksInWebParts.concat(linksInContent) // Merge Webparts and Main into a single sorted Array
  allLinks = Array.from((new Map(allLinks.map((item) => [item.join(), item]))).values());  // Remove duplicates
  allLinks.sort();                                                                         // Sort
  console.log("    Total links to sites:                           " + allLinks.length + "\n");

  // Get all Files

  fileContents = fs.readFileSync(sourceSitesAudit, "utf8"); // Read in the Files from the SITEAUDIT file
  allFiles = fileContents.split("\n");
  allFiles.forEach(stringToArray)
  allFiles = allFiles.filter(function (el) { return el != ""; })
  console.log("    Total sites stored:                              " + allFiles.length + "\n");

  // Get all Pages

  fileContents = fs.readFileSync(sourcePagesAudit, "utf8"); // Read in the Page from the PAGEAUDIT file
  allPages = fileContents.split("\n");
  allPages.forEach(stringToArray)
  allPages = allPages.filter(function (el) { return el != ""; })


  function stringToArray(sourceStr, index, arr) {
    sourceStr = sourceStr.replace(/["]+/g, '')
    resArr = sourceStr.split(" | ")

    //strip trailing "/"
    if (resArr[0].substring(resArr[0].length - 1, resArr[0].length) == "/") {
      resArr[0] = resArr[0].substring(0, resArr[0].length - 1);
      // console.log(resArr[0]);
    }

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

  var workbook = new excel.Workbook();

  var summarySheet = workbook.addWorksheet('Summary');
  summarySheet.columns = [
    { header: 'Summary',   key: 'summary',   width: 110 },
    { header: 'Count',     key: 'count',     width:  10 },
    { header: 'Published', key: 'published', width:  10 }
  ];

  var validLinksSheet = workbook.addWorksheet('Valid Links');
  validLinksSheet.columns = [
    { header: 'Site URL',                      key: 'siteurl',                    width: 70 },
    { header: 'Parent Page URL',               key: 'parentpageurl',              width: 70 },
    { header: 'Location in Parent Page',       key: 'location',                   width: 20 },
    { header: 'Parent Page Moderation Status', key: 'parentstatus',               width: 20 },  
    { header: 'Parent Page Published Version', key: 'parentpagepublishedversion', width: 10 },
    { header: 'File Type',                     key: 'filetype',                   width: 10 }
  ];

  var brokenLinksSheet = workbook.addWorksheet('Broken Links');
  brokenLinksSheet.columns = [
    { header: 'Broken Link',                   key: 'siteurl',                    width: 70 },
    { header: 'Parent Page URL',               key: 'parentpageurl',              width: 70 },
    { header: 'Location in Parent Page',       key: 'location',                   width: 20 },
    { header: 'Parent Page Moderation Status', key: 'parentpagestatus',           width: 20 },
    { header: 'Parent Page Published Version', key: 'parentpagepublishedversion', width: 10 },   
    { header: 'File Type',                     key: 'filetype',                   width: 10 }
  ];

  var validDocsSheet = workbook.addWorksheet('Linked Sites');
  validDocsSheet.columns = [
    { header: 'Site URL',      key: 'siteurl',          width: 180 },
    { header: 'No of links',   key: 'numberoflinks',    width:  15 }
  ];

  var unusedDocsSheet = workbook.addWorksheet('Unlinked Sites');
  unusedDocsSheet.columns = [
    { header: 'Site URL',    key: 'siteurl',  width: 180 }
  ];

  // Summary Text
  summarySheet.addRow(["Comparing all links in Content and Web Parts to Gateway Sites against all the Sites stored in Gateway"]).commit();
  summarySheet.addRow(["Sites with no content links to them can still be linked to from the site navigation"]).commit();
  summarySheet.addRow([" "]).commit();

  // Get all links that go to a file held in SharePoint
  var validLinks = compareArrays(allLinks, allFiles, "valid links");
  var pubPages = countPublishedPages(validLinks, 6);
  summarySheet.addRow(["All links going to a site:",validLinks.length, pubPages]).commit();
  console.log("    All links going to a site:                      " + validLinks.length + "  " + pubPages);
  validLinks.forEach(OutputValidLinks)

  // Get all broken links
  var brokenLinks = compareArrays(allLinks, allFiles, "broken links");
  var pubPages = countPublishedPages(brokenLinks, 6);
  summarySheet.addRow(["All broken links or links to unpublished sites: ",brokenLinks.length, pubPages]).commit();
  console.log("    All broken links or links to unpublished sites:  " + brokenLinks.length + "    " + pubPages);
  brokenLinks.forEach(OutputBrokenLinks);

  // Get all files with 1 or more link to them
  var validDocs = compareArrays(allFiles, allLinks, "linked sites");
  var pubPages = countPublishedPages(validDocs, 9);
  summarySheet.addRow(["All sites with content links to them:",validDocs.length, pubPages]).commit();
  console.log("    All sites with content links to them:            " + validDocs.length + "   " + pubPages);
  validDocs.forEach(OutputValidDocs)

  // Get all files with no links to them
  var unusedDocs = compareArrays(allFiles, allLinks, "unlinked sites");
  var pubPages = countPublishedPages(unusedDocs, 9);
  summarySheet.addRow(["All sites with no content links to them:",unusedDocs.length, pubPages]).commit();
  console.log("    All sites with no content links to them:         " + unusedDocs.length + "   " + pubPages);
  unusedDocs.forEach(OutputUnusedDocs);

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
    validLinksSheet.addRow([record[0], record[1], record[2], record[4], record[6], record[10]]).commit();
    // fs.appendFileSync(validLinksFile, record[0] + " | " + record[1] + " | " + record[2] + " | " + record[3]);
    // fs.appendFileSync(validLinksFile, "\n");
  };

  function OutputBrokenLinks(record) {
    brokenLinksSheet.addRow([record[0], record[1], record[2], record[4], record[6], record[10]]).commit(); 
    // fs.appendFileSync(brokenLinksFile, '<li><a href=' + record[0] + '>' + record[0] + '</a> in  ' + record[2] + '</li>');  
    // fs.appendFileSync(brokenLinksFile, record[0] + " | " + record[1] + " | " + record[2]);
    // fs.appendFileSync(brokenLinksFile, "\n");
  };

  function OutputValidDocs(record) {
    validDocsSheet.addRow([record[0], record[11]]).commit();
    // fs.appendFileSync(validDocsFile, record[0] + " | " + record[3]);
    // fs.appendFileSync(validDocsFile, "\n");
  };

  function OutputUnusedDocs(record) {
    unusedDocsSheet.addRow([record[0]]).commit();
    // fs.appendFileSync(unusedDocsFile, record[0] + " | " + record[3]);
    // fs.appendFileSync(unusedDocsFile, "\n");
  };

  function countPublishedPagesToPublishedContent(arrayA, parentPage, content){
    var pagesCount = 0;
    for (i = 0; i < arrayA.length; i++) {
      if ((arrayA[i][parentPage] != "New") && (arrayA[i][parentPage] != "Unpublished") && (arrayA[i][content] != "New") && (arrayA[i][content] != "Unpublished")){
        pagesCount++
      }
    }
    return pagesCount
  }

  function countPublishedPagesToUnpublishedContent(arrayA, parentPage, content){
    var pagesCount = 0;
    for (i = 0; i < arrayA.length; i++) {
      if ((arrayA[i][parentPage] != "New") && (arrayA[i][parentPage] != "Unpublished") &&
          ((arrayA[i][content] == "New") || (arrayA[i][content] == "Unpublished")) &&
          (arrayA[i][4] != "Pending" )){
        pagesCount++
      }
    }
    return pagesCount
  }

  function countPublishedPages(arrayA, column){
    var pagesCount = 0;
    for (i = 0; i < arrayA.length; i++) {
      if ((arrayA[i][column] != "New") && (arrayA[i][column] != "Unpublished")){
        pagesCount++
      }
    }
    return pagesCount
  }

  function compareArrays(arrayA, arrayB, option) {
    var res = []

    switch (option) {

      case "valid links":   // All links where there is an associated stored file
        for (i = 0; i < arrayA.length; i++) {
          found = false;
          for (j = 0; j < arrayB.length; j++) {
            if (arrayA[i][0].toLowerCase() === arrayB[j][0].toLowerCase()) {
              found = true
              arrayA[i][4] = arrayA[i][5]  // Copy parent page moderation status field (uploaded for lists in [5])
              arrayA[i][3] = arrayB[j][5]  // Cpoy across the file moderation status field
              arrayA[i][5] = arrayB[j][9]  // Cpoy across the file Published Version field
            }
          }
          if (found) {
            // arrayA[i] has an associated page
            // Get the page that the link is on and find it's status (arrayA[i][1])
            foundParent = false;
            for (k = 0; k < allPages.length; k++) {
              if (arrayA[i][1].toLowerCase() === allPages[k][0].toLowerCase()) {
                foundParent = true
                arrayA[i][4] = allPages[k][5]  // Cpoy across the parent page moderation status field
                arrayA[i][6] = allPages[k][9]  // Cpoy across the parent page Published Version field 
              }
            } 
            // If parentURL not found (a backend list) then set the field to ""
            // if (!foundParent){
            //   arrayA[i][4] = ""
            //}
            res.push(arrayA[i])
          }
        }
        return res

      case "broken links":  // All links where there isn't an associated page
        for (i = 0; i < arrayA.length; i++) {
          found = false;
          for (j = 0; j < arrayB.length; j++) {
            if (arrayA[i][0].toLowerCase() === arrayB[j][0].toLowerCase()) {
              found = true;
            }
          }
          if (!found) {
             // console.log(arrayA[i])
            // arrayA[i] does not have an associated page
            // Get the page that the link is on and find it's status (arrayA[i][1])
            for (k = 0; k < allPages.length; k++) {
              if (arrayA[i][1].toLowerCase() === allPages[k][0].toLowerCase()) {
                arrayA[i][4] = allPages[k][5]  // Cpoy across the parent page modertion status field
                arrayA[i][6] = allPages[k][9]  // Cpoy across the parent page Published Version field
             }
            }
            res.push(arrayA[i])
          }
        }
        return res

      case "linked sites":  // All documents with at least one link to them
        for (i = 0; i < arrayA.length; i++) {
          found = false;
          linkCount=0
          for (j = 0; j < arrayB.length; j++) {
            if (arrayA[i][0].toLowerCase() === arrayB[j][0].toLowerCase()) {
              found = true
              linkCount++
            }
          }
          if (found) {
            arrayA[i][11] = linkCount
            res.push(arrayA[i])
          }
        }
        return res

      case "unlinked sites":   // All documemnts that have no links to them
        for (i = 0; i < arrayA.length; i++) {
          found = false;
          for (j = 0; j < arrayB.length; j++) {
            if (arrayA[i][0].toLowerCase() === arrayB[j][0].toLowerCase()) {
              found = true;
            }
          }
          if (!found) {
            res.push(arrayA[i])
          }
        }
        return res

    }

  }
}