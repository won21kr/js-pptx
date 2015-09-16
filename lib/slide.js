var Shape = require('./shape');
var Chart = require('./chart')


//======================================================================================================================
// Slide
//======================================================================================================================

var Slide = function (args) {
  this.content = args.content;
  this.presentation = args.presentation;
  this.name = args.name;

  // TODO: Validate arguments
};

Slide.prototype.getShapes = function () {

  // TODO break out getShapeTree
  return this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].map(function (sp) {
    return new Shape(sp);
  });
};

Slide.prototype.addChart = function (data, done) {
  var self = this;
  var chartName = "chart"+(this.presentation.getChartCount() + 1);
  var chart = new Chart({slide: this, presentation: this.presentation, name: chartName});
  var slideName = this.name;

  chart.load(data, function(err, data) { // TODO pass it real data
    self.content["p:sld"]["p:cSld"][0]["p:spTree"][0]["p:graphicFrame"] = chart.content; //jsChartFrame["p:graphicFrame"];

    // Add entry to slide1.xml.rels
    // There should a slide-level and/or presentation-level method to add/track rels
    var rels = self.presentation.content['ppt/slides/_rels/' + slideName + '.xml.rels'];
    var numRels = rels["Relationships"]["Relationship"].length;
    var rId = "rId"+(numRels + 1);
    var numRels = rels["Relationships"]["Relationship"].push({
      "$": {
        "Id": rId ,
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        "Target": "../charts/" + chartName + ".xml"
      }
    });
    done(null, self);
  });
};

Slide.prototype.addShape = function () {
  var shape = new Shape();
  this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].push(shape.content);
  return shape;
};

module.exports = Slide;