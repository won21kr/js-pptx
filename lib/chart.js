var fragment = require('./fragment');

var Chart = module.exports = function (args) {
  this.presentation = args.presentation;  // should this be inferred from the slide?
  this.slide = args.slide;
  this.name = args.name;
}

// TODO  Generate Excel Worksheet for dataseries
// TODO  Generate data series reference for chart xml


Chart.prototype.load = function (data, done) {
  var self = this;
  var chartName = this.name; //'chart1';
  var presentation = this.presentation;
  var worksheetName = 'Microsoft_Excel_Sheet1.xlsx'

  // 'ppt/charts/chart1.xml'
  var jsChart = require('./fragments/js/chart');
  var jsChartSeries = require('./fragments/js/chartseries');

  // TODO Generate the chart series from the data
  jsChart["c:chartSpace"]["c:chart"][0]["c:plotArea"][0]["c:barChart"][0]["c:ser"] = jsChartSeries["c:ser"];

  presentation.content['ppt/charts/chart1.xml'] = jsChart;

  // '[Content_Types].xml' .. add references to the chart and the spreadsheet
  presentation.content["[Content_Types].xml"]["Types"]["Override"].push({
    "$": {
      "PartName": "/ppt/charts/" + chartName + ".xml",
      "ContentType": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    }
  });
  presentation.content["[Content_Types].xml"]["Types"]["Default"].push({
    "$": {
      "Extension": "xlsx",
      "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
  });


  // embeddings/Microsoft_Excel_Sheet1.xlsx
  // TODO: Generate the Excel from the chart data
  fragment.fromBinary(worksheetName, function (err, data) {
    presentation.content["ppt/embeddings/" + worksheetName] = data;

    // ppt/charts/_rels/chart1.xml.rels
    // TODO: Don't assume we need to create it, read it if it exists and increment the rID
    presentation.content["ppt/charts/_rels/" + chartName + ".xml.rels"] = {

      "Relationships": {
        "$": {
          "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
        },
        "Relationship": [
          {
            "$": {
              "Id": "rId1",
              "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
              "Target": "../embeddings/" + worksheetName
            }
          }
        ]
      }
    };

    var jsChartFrame = require('./fragments/js/chartframe');
    self.content = jsChartFrame["p:graphicFrame"];
    done(null, self);

  });
}