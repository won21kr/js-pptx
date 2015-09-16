var excelbuilder = require('./msexcel-builder');
var fs = require('fs');
var clone = require('./util/clone')


var Chart = module.exports = function (args) {
  this.presentation = args.presentation;  // should this be inferred from the slide?
  this.slide = args.slide;
  this.name = args.name;
}

var XmlNode = require('./xmlnode');

// TODO  Generate Excel Worksheet for dataseries
// TODO  Generate data series reference for chart xml


Chart.prototype.load = function (chartInfo, done) {
  var self = this;
  var chartName = this.name; // e.g. 'chart1';

  var jsChartFrame = clone(require('./fragments/js/chartframe'));
  jsChartFrame["p:graphicFrame"]["p:nvGraphicFramePr"]["p:nvPr"]["p:extLst"]["p:ext"]["p14:modId"] =
  {
    "$": {
      "xmlns:p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
      "val": Math.floor(Math.random() * 4294967295)
    }
  }
  self.content = jsChartFrame["p:graphicFrame"];


  // 'ppt/charts/chartN.xml'
  var chartContent = this.getChartBase(chartInfo);
  this.presentation.registerChart(chartName, chartContent);


  this.createWorkbook(chartInfo, function (err, workbookJSZip) {
    self.presentation.registerChartWorkbook(chartName,  workbookJSZip.generate({type: 'arraybuffer'}));
    done(null, self);
  });
}

Chart.prototype.getChartBase = function (chartInfo) {
  var jsChart = clone(require('./fragments/js/chart'));
  var jsChartSeries = this.getChartSeries(chartInfo);

  // TODO Generate the chart series from the data
  jsChart["c:chartSpace"]["c:chart"][0]["c:plotArea"][0]["c:barChart"][0]["c:ser"] = jsChartSeries['c:ser'];
  return jsChart;
};

Chart.prototype.getChartSeries = function (chartInfo) {
  var res = {
    "c:ser": chartInfo.data.map(this._ser, this)
  }
  return res;
};


/// @brief returns XML snippet for a chart dataseries
Chart.prototype._ser = function (serie, i) {
  var rc2a = this._rowColToSheetAddress; // shortcut


  var ser = XmlNode()
      .addChild('c:idx', XmlNode().attr('val', i))
      .addChild('c:order', XmlNode().attr('val', i))
      .addChild('c:tx', this._strRef('Sheet1!' + rc2a(1, 2 + i, true, true), [serie.name]))
      .addChild('c:invertIfNegative', XmlNode().attr('val', 0))
      .addChild('c:cat', this._strRef('Sheet1!' + rc2a(2, 1, true, true) + ':' + rc2a(2 + serie.labels.length - 1, 1, true, true), serie.labels))
      .addChild('c:val', this._numRef('Sheet1!' + rc2a(2, 2 + i, true, true) + ':' + rc2a(2 + serie.labels.length - 1, 2 + i, true, true), serie.values, "General"))

  if (serie.color) {
    ser.addChild('c:spPr',
        XmlNode().addChild('a:solidFill',
            XmlNode().addChild('a:srgbClr',
                XmlNode().attr('val', serie.color)
            )
        )
    );
  }
  else if (serie.schemeColor) {
    ser.addChild('c:spPr',
        XmlNode().addChild('a:solidFill',
            XmlNode().addChild('a:schemeClr',
                XmlNode().attr('val', serie.schemeColor)
            )
        )
    );
  }

  return ser.el;
};


///
/// @brief Transform an array of string into an office's compliance structure
///
/// @param[in] region String
///		The reference cell of the string, for example: $A$1
/// @param[in] stringArr
///		An array of string, for example: ['foo', 'bar']
///
Chart.prototype._strRef = function (region, stringArr) {

  var strRef = XmlNode().addChild('c:strRef', XmlNode()
      .addChild('c:f', region)
      .addChild('c:strCache', this._strCache(stringArr))
  );
  return strRef.el;
}


Chart.prototype._strCache = function (stringArr) {
  var strRef = XmlNode().addChild('c:ptCount', XmlNode().attr('val', stringArr.length))
  for (var i = 0; i < stringArr.length; i++) {
    strRef.addChild('c:pt', XmlNode().attr('idx', i).addChild('c:v', stringArr[i]))
  }

  return strRef;
}


///
/// @brief Transform an array of numbers into an office's compliance structure
///
/// @param[in] region String
///		The reference cell of the string, for example: $A$1
/// @param[in] numArr
///		An array of numArr, for example: [4, 7, 8]
/// @param[in] formatCode
///		A string describe the number's format. Example: General
///
Chart.prototype._numRef = function (region, numArr, formatCode) {

  var numCache = XmlNode()
      .addChild('c:formatCode', formatCode)
      .addChild('c:ptCount', XmlNode().attr('val', numArr.length));

  for (var i = 0; i < numArr.length; i++) {
    numCache.addChild('c:pt', XmlNode().attr('idx', i).addChild('c:v', numArr[i].toString()));
  }

  var numRef = XmlNode().addChild('c:numRef', XmlNode()
      .addChild('c:f', region)
      .addChild('c:numCache', numCache)
  );
  return numRef.el;

}


Chart.prototype._rowColToSheetAddress = function (row, col, isRowAbsolute, isColAbsolute) {
  var address = "";

  if (isColAbsolute)
    address += '$';

  // these lines of code will transform the number 1-26 into A->Z
  // used in excel's cell's coordination
  while (col > 0) {
    var num = col % 26;
    col = (col - num ) / 26;
    address += String.fromCharCode(65 + num - 1);
  }

  if (isRowAbsolute)
    address += '$';

  address += row;

  return address;
};


// takes an  array with series data
// callback takes two parameters:
//    @err   Error, null if successful
//    @wb    JSZip object containing the workbook
Chart.prototype.createWorkbook = function (data, callback) {

  // First, generate a temporatory excel file for storing the chart's data
  var workbook = excelbuilder.createWorkbook();

  // Create a new worksheet with 10 columns and 12 rows
  // number of columns: data['data'].length+1 -> equaly number of series
  // number of rows: data['data'][0].values.length+1
  var sheet1 = workbook.createSheet('Sheet1', data['data'].length + 1, data['data'][0].values.length + 1);
  var headerrow = 1;

  // write header using serie name
  for (var j = 0; j < data['data'].length; j++) {
    sheet1.set(j + 2, headerrow, data['data'][j].name);
  }

  // write category column in the first column
  for (var j = 0; j < data['data'][0].labels.length; j++) {
    sheet1.set(1, j + 2, data['data'][0].labels[j]);
  }

  // for each serie, write out values in its row
  for (var i = 0; i < data['data'].length; i++) {
    for (var j = 0; j < data['data'][i].values.length; j++) {
      // col i+2
      // row j+1
      sheet1.set(i + 2, j + 2, data['data'][i].values[j]);
    }
  }
  workbook.generate(callback); // returns (err, zip)
};