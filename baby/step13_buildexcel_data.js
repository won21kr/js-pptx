var excelbuilder = require('msexcel-builder');

// Create a new workbook file in current working-path


var old = function(data, callback) {


  // Create a new worksheet with 10 columns and 12 rows
  var workbook = excelbuilder.createWorkbook('./', 'sample.xlsx')
  var sheet1 = workbook.createSheet('sheet1', 10, 12);

  // Fill some data
  sheet1.set(1, 1, 'I am title');
  for (var i = 2; i < 5; i++)
    sheet1.set(i, 1, 'test'+i);

  // Save it
  workbook.save(callback);
};


var GLOBAL_CHART_COUNT = 0;
createWorkbook = function(data, callback) {
  GLOBAL_CHART_COUNT += 1;
  var tmpExcelFile = 'sample'+GLOBAL_CHART_COUNT+'.xlsx';


  // First, generate a temporatory excel file for storing the chart's data
  var workbook = excelbuilder.createWorkbook('./', tmpExcelFile);

  // Create a new worksheet with 10 columns and 12 rows
  // number of columns: data['data'].length+1 -> equaly number of series
  // number of rows: data['data'][0].values.length+1
  var sheet1 = workbook.createSheet('Sheet1', data['data'].length+1, data['data'][0].values.length+1);
  var headerrow = 1;
  console.log("STARTED WORKBOOK...");

  // write header using serie name
  for( var j=0; j < data['data'].length; j++ ) {
    sheet1.set(j+2, headerrow, data['data'][j].name);
  }

  // write category column in the first column
  for( var j=0; j < data['data'][0].labels.length; j++ ) {
    sheet1.set(1, j+2, data['data'][0].labels[j]);
  }

  // for each serie, write out values in its row
  for (var i = 0; i < data['data'].length; i++) {
    for( var j=0; j < data['data'][i].values.length; j++ )
    {
      // col i+2
      // row j+1
      sheet1.set(i+2, j+2, data['data'][i].values[j]);
    }
  }
  console.log("FILLING WORKBOOK...");
  // Fill some data
  // Save it


  console.log("SAVING WORKBOOK...");
  workbook.save(function(err){
    if (err) {
      workbook.cancel();
      callback(err);
    }
    else  {
      console.log("SAVED: " + workbook.length);
      callback(null, workbook);
    }
  });
};


var barChart = {
  title: 'Sample bar chart',
  renderType: 'bar',
  xmlOptions: {
    "c:title": {
      "c:tx": {
        "c:rich": {
          "a:p": {
            "a:r": {
              "a:t": "Override title via XML"
            }
          }
        }
      }
    }
  },
  data: [
    {
      name: 'europe',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.5, 2.6, 2.8],
      color: 'ff0000'
    },
    {
      name: 'namerica',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.5, 2.7, 2.9],
      color: '00ff00'
    },
    {
      name: 'asia',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.1, 2.2, 2.4],
      color: '0000ff'
    },
    {
      name: 'lamerica',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.3, 0.3, 0.3],
      color: 'ffff00'
    },
    {
      name: 'meast',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.2, 0.3, 0.3],
      color: 'ff00ff'
    },
    {
      name: 'africa',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.1, 0.1, 0.1],
      color: '00ffff'
    }

  ]
};

createWorkbook(barChart, function(err, workbook) {console.log("ALL DONE"); });

//old(barChart, function(err, workbook) {console.log("ALL DONE"); });