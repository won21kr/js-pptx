var Shape = require('./shape');

//======================================================================================================================
// Slide
//======================================================================================================================

var Slide = function (content) {
  this.content = content;
}

Slide.prototype.getShapes = function () {
  return this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].map(function (sp) {
    return new Shape(sp);
  });
}

Slide.prototype.addShape = function() {
  var shape = new Shape();
  this.content["p:sld"]["p:cSld"][0]["p:spTree"][0]['p:sp'].push(shape.content);
  return shape;
}

module.exports = Slide;