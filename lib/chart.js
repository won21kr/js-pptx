var fragment = require('./fragment');
var excelbuilder = require('./msexcel-builder');
var fs =require('fs');

var Chart = module.exports = function (args) {
  this.presentation = args.presentation;  // should this be inferred from the slide?
  this.slide = args.slide;
  this.name = args.name;
}

var XmlNode = require('./xmlnode');

// TODO  Generate Excel Worksheet for dataseries
// TODO  Generate data series reference for chart xml


Chart.prototype.load = function (data, done) {
  var self = this;
  var chartName = this.name; // e.g. 'chart1';
  var presentation = this.presentation;
  var worksheetName = 'Microsoft_Excel_Sheet1.xlsx'

  // 'ppt/charts/chart1.xml'
  var jsChart = require('./fragments/js/chart');
  var jsChartSeries = require('./fragments/js/chartseries');

  // TODO Generate the chart series from the data
  jsChart["c:chartSpace"]["c:chart"][0]["c:plotArea"][0]["c:barChart"][0]["c:ser"] = jsChartSeries["c:ser"];

  presentation.content['ppt/charts/chart1.xml'] = jsChart;

  // '[Content_Types].xml' .. add references to the chart and the spreadsheet
  presentation.content["[Content_Types].xml"]["Types"]["Override"].push(XmlNode()
      .attr('PartName', "/ppt/charts/" + chartName + ".xml")
      .attr('ContentType', "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
      .el
  );

  presentation.content["[Content_Types].xml"]["Types"]["Default"].push(XmlNode()
      .attr('Extension', 'xlsx')
      .attr('ContentType', "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .el
  );


  // embeddings/Microsoft_Excel_Sheet1.xlsx

  this.createWorkbook(data, function (err, data) {

    presentation.content["ppt/embeddings/" + worksheetName] = data.generate({type: 'arraybuffer'});

    // ppt/charts/_rels/chart1.xml.rels
    presentation.content["ppt/charts/_rels/" + chartName + ".xml.rels"] = XmlNode().setChild("Relationships", XmlNode()
        .attr({
          'xmlns': "http://schemas.openxmlformats.org/package/2006/relationships"
        })
        .addChild('Relationship', XmlNode().attr({
          "Id": "rId1",
          "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
          "Target": "../embeddings/" + worksheetName
        }))).el


    var jsChartFrame = require('./fragments/js/chartframe');
    self.content = jsChartFrame["p:graphicFrame"];
    done(null, self);

  });
}

Chart.prototype.createWorkbook = function(data, callback) {

  // First, generate a temporatory excel file for storing the chart's data
  var workbook = excelbuilder.createWorkbook();

    // Create a new worksheet with 10 columns and 12 rows
  // number of columns: data['data'].length+1 -> equaly number of series
  // number of rows: data['data'][0].values.length+1
  var sheet1 = workbook.createSheet('Sheet1', data['data'].length+1, data['data'][0].values.length+1);
  var headerrow = 1;

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
  workbook.generate(callback); // returns (err, zip)
};