var XmlNode = function(tag, el) {
  this.el = el;
  this.tag = tag;
}

module.exports = XmlNode;

XmlNode.prototype.toJSON = function() { return this.el; }

XmlNode.prototype.attr = function(key, val) {
  if (typeof key === 'undefined') {
    return this.el['$'];
  }
  else if (typeof val === 'undefined') {
    return this.el['$'] ? this.el['$'][key] : undefined;
  }
  else {
    if (!this.el['$']) this.el['$'] = {};
    this.el['$'][key] = val;
  }
}

XmlNode.prototype.getChildren = function(tag) {


}