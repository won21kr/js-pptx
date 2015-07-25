var Shape = require('./shape');
var OfficeChart = require('./officechart');

var fragment = require('./fragment');


//======================================================================================================================
// Slide
//======================================================================================================================

var Slide = function (content, presentation) {
  this.content = content;
  this.presentation = presentation;
}

Slide.prototype.getShapes = function () {
  return this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].map(function (sp) {
    return new Shape(sp);
  });
}

Slide.prototype.addChart = function (done) {
  var chartName = 'chart1';
  var slideName = 'slide1';
  var presentation = this.presentation;
  var worksheetName = 'Microsoft_Excel_Sheet1.xlsx'
  var slide = this;

  fragment.fromXml('chart.html', function (err, jsChart) {

    // 'ppt/charts/chart1.xml'
    presentation.content['ppt/charts/chart1.xml'] = jsChart;

    // '[Content_Types].xml'
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


    //<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>


    // embeddings/Microsoft_Excel_Sheet1.xlsx
    fragment.fromBinary(worksheetName, function (err, data) {
      presentation.content["ppt/embeddings/" + worksheetName] = data;
      console.log(data.length);



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
      }



      // add the chart frame to the slide.xml
      fragment.fromXml('slide_chart_frame.xml', function(err, jsChartFrame) {
//        console.log(JSON.stringify(jsChartFrame, null,4))
        console.log(JSON.stringify(slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0], null,4))
        slide.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"] = jsChartFrame["p:graphicFrame"];

        // slide1.xml.rels
        //console.log(Object.keys(presentation.content))

        var rels = presentation.content['ppt/slides/_rels/' + slideName + '.xml.rels'];
        var numRels = rels["Relationships"]["Relationship"].length;
        var rId = "rId"+(numRels + 1);
        var numRels = rels["Relationships"]["Relationship"].push({
          "$": {
            "Id": rId ,
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
            "Target": "../charts/chart1.xml"
          }
        });

        done(null, jsChart);

      })


    });

  });

//
//  this.presentation.content['ppt/charts/chart1.xml']
}

Slide.prototype.addShape = function () {
  var shape = new Shape();
  this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].push(shape.content);
  return shape;
}

module.exports = Slide;