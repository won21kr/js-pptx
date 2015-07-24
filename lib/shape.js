var spPr = require('./spPr');
var defaults = require('./defaults');
function clone(obj) { return JSON.parse(JSON.stringify(obj))}

//======================================================================================================================
// Shape
//======================================================================================================================

var Shape = function (content) {

  this.content = content || clone(defaults["p:sp"]);
}

Shape.prototype.text = function (text) {
  if (text) {
    this.content['p:txBody'][0]['a:p'][0]['a:r'][0]['a:t'][0] = text;
    return this;
  }
  else {
    return this.content['p:txBody'][0]['a:p'][0]['a:r'][0]['a:t'][0];
  }
}

Shape.prototype.shapeProperties = function () {
  return new spPr(this.content["p:spPr"][0])
}

module.exports = Shape;