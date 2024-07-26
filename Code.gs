

function createSheetWithFileIds(fileIds) {

  var data = [['File Name', 'File ID']];

  fileIds.forEach(function(id) {
    var file = DriveApp.getFileById(id);
    data.push([file.getName(), file.getId()]);
  });


  var sheetName = 'Analyzed Data';  

var sheets = createSheet(sheetName);

         

         






var csvData;
for (let i=0; i<data.length; i++)
{
 csvData += loadCsvFromDrive(data[i][1]);
}
  

  const colNum_seller = getColNumsByStr('SellerName', csvData);
  const colNum_kpi = getColNumsByStr('Kpi', csvData);


 
        t0 = [];
        for (let j=0; j<colNum_kpi.length; j++)
        {
          t0.push([]);
        }


sellerKpis = new Map();
        
        for (let j=0; j<colNum_kpi.length; j++)
        {
            kpiArr = [];
            for (let i=1; i<csvData.length; i++)
            {
                
                const seller = csvData[i][colNum_seller[0]];

                if (!sellerKpis.has(seller))
                {
                    sellerKpis.set(seller, deepCopy2DArray(t0));
                }
        
                
                const kpi = csvData[i][colNum_kpi[j]];
                        
        
                kpiArr = sellerKpis.get(seller);
                kpiArr[j].push(kpi);
                sellerKpis.set(seller, kpiArr);
            }
            
        }
        
        
       sellerKpis.forEach(
            (value, key) =>
            {
                for (let j=0; j<value.length; j++)
                {
                ans = []; 
                a = 0;

                for (let i=0; i<value[j].length; i++)
                {      
                    const v = parseInt(value[j][i], 10);
                    if (v < 4)
                    {                 
                        a += v;
                        
                    }
                    r = a / (5 * value[j].length);
                    
                }
                ans.push(r);
                value[j] = ans;
                sellerKpis.set(key, value);
                }
            }
        );
        

sellerKpis.forEach(
            (value, key) =>
            {
                                t = []; 
                a = 0;

                for (let j=0; j<value.length; j++)
                {

                for (let i=0; i<value[j].length; i++)
                {      
                    let v = parseFloat(value[j][i]);
                    a += v;
                    
                    
                }
                }

                r_mean = a / value.length;

                t.push(r_mean);
                value.push(t);
                sellerKpis.set(key, value);
                


            }
        );
        




sellerKpis.forEach(
            (value, key) =>
            {
                                t = []; 

                for (let j=0; j<value.length; j++)
                {

                for (let i=0; i<value[j].length; i++)
                {      
                    let v = parseFloat(value[j][i]);
                    t.push(v);
                    
                    
                }
                }

                value = t;
                sellerKpis.set(key, value);
                


            }
        );
        



riskCsv = []; riskNameArr = [];

riskNameArr.push('Seller');        

for (let j=0; j<colNum_kpi.length; j++)
{
    const kpiName = csvData[0][colNum_kpi[j]];
    const pattern = /Kpi/g;
    const riskName = kpiName.replace(pattern, 'Risk');
    riskNameArr.push(riskName);        
    
}
riskNameArr.push('MeanRisk'); 
riskCsv.push(riskNameArr); 




sellerKpis.forEach(
            (value, key) =>
            {
                t = [];
                t.push(key);
                for (let i=0; i<value.length; i++)
                  t.push(value[i]);
                riskCsv.push(t);

            }
        );
 


var sheet = sheets.getSheets()[0];  

    
    pasteCsvData(sheet, csvData, 1);
   
   
       sheets.insertSheet('Risks');
  var sheet2 = sheets.getSheets()[1];
    pasteCsvData(sheet2, riskCsv, 1);
    
    plotBarGraph1(sheet2,'RiskPunctual, RiskProdQuality, MeanRisk of sellers');


unitPriceIndex = getColNumsByStr('Price', csvData)[0];
sellerPrices = new Map();

            for (let i=1; i<csvData.length; i++)
            {
                
                const seller = csvData[i][colNum_seller[0]];

                if (!sellerPrices.has(seller))
                {
                    sellerPrices.set(seller, []);
                }

s = sellerPrices.get(seller);
s.push(csvData[i][unitPriceIndex]);
                sellerPrices.set(seller, s);
        
                      
            }

          
       sellerPrices.forEach(
            (value, key) =>
            {
              totalProd = value.length;
              a = 0;
                for (let i=0; i<totalProd; i++)
                {
                  a += parseFloat(value[i]);
                }

                m = a / totalProd;

                sellerPrices.set(key, m);
 
            }
        );
        

 

priceCsv = []; riskNameArr = [];

riskNameArr.push('Seller');
riskNameArr.push('MeanPrice');        
priceCsv.push(riskNameArr); 




sellerPrices.forEach(
            (value, key) =>
            {
                t = [];
                t.push(key);
                t.push(value);

                priceCsv.push(t);

            }
        );        




       sheets.insertSheet('Prices');
var sheet3 = sheets.getSheets()[2];
    pasteCsvData(sheet3, priceCsv, 1);

plotBarGraph(sheet3, 'Mean Price of Sellers');



const expDateIndex = getColNumsByStr('Expiry', csvData)[0];
const prodIndex = getColNumsByStr('Prod', csvData)[0];
const contactIndex = getColNumsByStr('Contact', csvData)[0];

ma = [];

for (let i = 1; i<csvData.length; i++)
{
    e = csvData[i][expDateIndex];
    p = csvData[i][prodIndex];
    c = csvData[i][contactIndex];
    
    if (checkTimestamp(e))
    {
        ma.push([p,e,c]);
    }
        
}

for (let i = 0; i<ma.length; i++)
{
    sendEmail(ma[i]);
}



const tsI = getColNumsByStr('DateP', csvData)[0];
const stI = getColNumsByStr('SubTotal', csvData)[0];

yearTotal = new Map();

for (let i=1; i<csvData.length; i++)
{
  ss = '';
    t = csvData[i][tsI];
    for (j=0; j<4; j++)
    {
        ss += t[j];
    }
    y = parseFloat(ss);
    if (!yearTotal.has(y))
    {
        yearTotal.set(y, 0);
    }

    a = yearTotal.get(y) + parseFloat(csvData[i][stI]); 
    yearTotal.set(y, a);
        

}



const sortedEntries = Array.from(yearTotal.entries()).sort((a, b) => a[0] - b[0]);


yearTotal = new Map(sortedEntries);


 

yearCsv = []; riskNameArr = [];

riskNameArr.push('Year');
riskNameArr.push('TotalSpent');        
yearCsv.push(riskNameArr); 




yearTotal.forEach(
            (value, key) =>
            {
                t = [];
                t.push(key);
                t.push(value);

                yearCsv.push(t);

            }
        );        




       sheets.insertSheet('Spendings');
var sheet4 = sheets.getSheets()[3];
    pasteCsvData(sheet4, yearCsv, 1);

plotGraph(sheet4,'Total Spent vs Time');








k_m = [];
for (let i=1; i<csvData.length; i++)
{
    ps = 0;
    for (let j=0; j<colNum_kpi.length; j++)
    {
     ps += parseFloat(csvData[i][colNum_kpi[j]]);
    }
    k_m.push(ps / 2);
}

s = [];
for (let i=1; i<csvData.length; i++)
{
    s.push(csvData[i][colNum_seller[0]]);
}


sk = new Map();

for (let i=0; i<k_m.length; i++)
{
    si = s[i];

                if (!sk.has(s[i]))
                {
                    sk.set(s[i], []);
                }

    t = sk.get(si);
    t.push(k_m[i]);
    sk.set(si, t);
}

          
       sk.forEach(
            (v, k) =>
            {
                su = 0;
                for (let i=0; i<v.length; i++)
                {
                    su += v[i];
                }
                m = su / v.length;
                sk.set(k, m);
                

 
            }
        );   


m2 = [...sk.values()];
b = 0;
a = -999999999;
for (let c=0; c<m2.length; c++)
{
    v=m2[c];
    if (v > a)
    {
        a = v;
        b = c;
    }
}

m1 = [...sk.keys()];
bestSeller = m1[b];



perfCsv = []; riskNameArr = [];

riskNameArr.push('Seller');
riskNameArr.push('OverallPerformance');        
perfCsv.push(riskNameArr); 




sk.forEach(
            (value, key) =>
            {
                t = [];
                t.push(key);
                t.push(value);

                perfCsv.push(t);

            }
        );        




       sheets.insertSheet('Recommendation');
var sheet5 = sheets.getSheets()[4];
    pasteCsvData(sheet5, perfCsv, 1);

plotBarGraph(sheet5, 'Overall Performance of Sellers');




  sheet5.getRange("C9").setValue(`The most reliable seller is:  
  ${bestSeller}`);


return sheets.getUrl();


}


function getColNumsByStr(substr, csvData)
{
    let ans = []
    for (let i=0; i < csvData[0].length; i++)
    {
        if (csvData[0][i].includes(substr))
        {
           ans.push(csvData[0].indexOf(csvData[0][i]));
        }
    }
    return ans;
}













function loadCsvFromDrive(fileId) {
  var file = DriveApp.getFileById(fileId);
  var csvContent = file.getBlob().getDataAsString();
  var csvData = Utilities.parseCsv(csvContent);
  return csvData; 
}



function createSheet(sheetName) {
  
  var scriptFileId = DriveApp.getFileById(ScriptApp.getScriptId()).getId();
  
  
  var currentFolder = DriveApp.getFileById(scriptFileId).getParents().next();
  
  
  var newSpreadsheet = SpreadsheetApp.create(sheetName);
  
  
  var newFile = DriveApp.getFileById(newSpreadsheet.getId());
  currentFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);  


  return newSpreadsheet;
}




function pasteCsvData(sheet, csvDat3a, startRow) {
  sheet.getRange(startRow, 1, csvDat3a.length, csvDat3a[0].length).setValues(csvDat3a);
}




function plotGraph(sheet, title) {
  var range =  sheet.getDataRange();
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(2, 5, 0, 0)
    .setOption('title', title)
    .build();
  sheet.insertChart(chart);
}


function plotBarGraph(sheet, title) {
  var range =  sheet.getDataRange();
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(range)
    .setPosition(2, 5, 0, 0)
    .setOption('title', title)
    .build();
  
  sheet.insertChart(chart);
}




function sendEmail(ma) {
  var emailAddress = Session.getActiveUser().getEmail(); 

  p = ma[0]; e = ma[1]; c = ma[2];
  
  var subject = 'Expiry Notification for ' + p;
  var body =
  `Dear Customer,

  Your product is expiring soon. Here are the details:

  Product Name: ${p}
  Expiry Date: ${e}


  You may renew in just a few clicks below.
  Seller: ${c}`
  

  MailApp.sendEmail(emailAddress, subject, body);
}




function convertToTimestamp(dateString) {
    return new Date(dateString).getTime();
}


function getCurrentTimestamp() {
    return Date.now();
}


const threshold = 5 * 24 * 60 * 60 * 1000;


function checkTimestamp(dateString) {
    const c = convertToTimestamp(dateString); 
    const u = getCurrentTimestamp(); 

if (c > u)
{
    if (c - u < threshold) {
        return 1;
    }
}
    return 0;
}




        function deepCopy2DArray(array) {
  return array.map(row => row.slice());
}

function onOpen() {
  var card = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('File ID Collector'))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextButton()
          .setText('Select Files and Generate Sheet')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('showSidebar'))))
      .build();
  return card;
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Select Files');
  return html;
}






function plotBarGraph1(sheet, title) {
  var range = sheet.getRange("A1:D" + sheet.getLastRow());
  
  
  var data = range.getValues();
  
  
  var xValues = [];
  var yValues = [];
  for (var i = 1; i < data.length; i++) {
    xValues.push([data[i][0]]); 
    yValues.push([data[i][3]]); 
  }
  
  
  var chartData = [];
  for (var j = 0; j < xValues.length; j++) {
    chartData.push([xValues[j][0], yValues[j][0]]);
  }
  
  
  var chart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange("A1:D" + sheet.getLastRow()))
      .setOption('useFirstColumnAsDomain', true)
      .setOption('series', {
        0: {targetAxisIndex: 0}
      })
      .setPosition(3, 5, 0, 0)
      .setOption('title', title)
      .build();
  
  sheet.insertChart(chart);
}
