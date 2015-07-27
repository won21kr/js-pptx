var fragment = require('./fragment');
var excelbuilder = require('./msexcel-builder');
var fs = require('fs');

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
  var presentation = this.presentation;
  var worksheetName = 'Microsoft_Excel_Sheet1.xlsx'

  // 'ppt/charts/chart1.xml'
  var content = this.getChartBase(chartInfo);

  presentation.content['ppt/charts/chart1.xml'] = content;

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

  this.createWorkbook(chartInfo, function (err, data) {

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

Chart.prototype.getChartBase = function (chartInfo) {
  var jsChart = require('./fragments/js/chart');
  var jsChartSeries = this.getChartSeries(chartInfo);

  // TODO Generate the chart series from the data
  jsChart["c:chartSpace"]["c:chart"][0]["c:plotArea"][0]["c:barChart"][0]["c:ser"] = jsChartSeries;
  return jsChart;
};

Chart.prototype.getChartSeries = function (chartInfo) {
//  var res =require('./fragments/js/chartseries');
//
//  return res;
  var res = {
    "c:ser": chartInfo.data.map(this._ser, this)
  }
//  console.json(res);
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
  console.json(ser)

  return ser.el;
//
//  if (serie.color) {
//    serieData['c:ser']['c:spPr'] = {
//      'a:solidFill': {
//        'a:srgbClr': {'@val': serie.color}
//      }
//    };
//  }
//  else if (serie.schemeColor) {
//    serieData['c:ser']['c:spPr'] = {
//      'a:solidFill': {
//        'a:schemeClr': {'@val': serie.schemeColor}
//      }
//    };
//  }
//
//  if (serie.xml) {
//    serieData['c:ser'] = _.merge(serieData['c:ser'], serie.xml)
//  }
//
//
//  // for pie charts
//  if (serie.colors) {
//    serieData['c:ser']['#list'] = this._colorRef(serie.colors);
//  }


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
  //console.log(JSON.stringify(numRef.el,null,4))
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
},


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