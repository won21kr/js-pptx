var defaults = require('./defaults');
var XmlNode = require('./xmlnode');
function clone(obj) { return JSON.parse(JSON.stringify(obj))}

//======================================================================================================================
// spPr   ShapeProperties
//======================================================================================================================

var shapeProperties = function (content) {
  this.content = content || clone(defaults["p:sp"]["p:spPr"][0]);
}

shapeProperties.prototype.toJSON = function() {
  return {
    x: this.x(),
    y: this.y(),
    cx: this.cx(),
    cy: this.cy(),
    prstGeom: this.prstGeom()
  }
}

shapeProperties.prototype.x = function(val) {
  if (arguments.length == 0) return this.content["a:xfrm"][0]["a:off"][0]['$'].x;
  else  this.content["a:xfrm"][0]["a:off"][0]['$'].x = val;
  return this;
}
shapeProperties.prototype.y = function(val) {
  if (arguments.length == 0) return this.content["a:xfrm"][0]["a:off"][0]['$'].y;
  else this.content["a:xfrm"][0]["a:off"][0]['$'].y = val;
  return this;
}
shapeProperties.prototype.cx = function(val) {
  if (arguments.length == 0) return this.content["a:xfrm"][0]["a:ext"][0]['$'].cx;
  else this.content["a:xfrm"][0]["a:ext"][0]['$'].cx = val;
  return this;
}
shapeProperties.prototype.cy = function(val) {
  if (arguments.length == 0) return this.content["a:xfrm"][0]["a:ext"][0]['$'].cy;
  else this.content["a:xfrm"][0]["a:ext"][0]['$'].cy = val;
  return this;
}

// see http://www.officeopenxml.com/drwSp-prstGeom.php
shapeProperties.prototype.prstGeom = function(shape) {
  if (arguments.length == 0) return this.content["a:prstGeom"][0]["$"]["prst"] ;
  else this.content["a:prstGeom"][0]["$"]["prst"] = shape;
  return this;
}

module.exports = shapeProperties;

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.presetgeometry(v=office.14).aspx