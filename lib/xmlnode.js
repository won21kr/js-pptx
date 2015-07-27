var XmlNode = function (el) {
  if (this instanceof XmlNode) {
    this.el = el || {};
  }
  else return new XmlNode(el);
}

module.exports = XmlNode;

XmlNode.prototype.toJSON = function () {
  return this.el;
}

XmlNode.prototype.attr = function (key, val) {
  if (typeof val == 'number') val = ''+val;
  if (typeof key === 'undefined') {
    // return all attributes, if any are defined
    return this.el['$'];
  }
  else if (arguments.length == 1 && (typeof key == 'object')) {
    // assign attributes
    this.el['$'] = this.el['$'] || {};
    for (var attrName in key) {
      if (key.hasOwnProperty(attrName)) {
        this.el['$'][attrName] = key[attrName];
      }
    }
    return this;
  }
  else if (typeof val === 'undefined') {
    // get the attribute value
    return this.el['$'] ? this.el['$'][key] : undefined;

  }
  else {
    this.el['$'] = this.el['$'] || {};
    this.el['$'][key] = val;
    return this;
  }
}

XmlNode.prototype.addChild = function (tag, node) {
//  if (typeof node == 'string') this.el[tag] = node;
//  else {
    this.el[tag] = this.el[tag] || [];
    this.el[tag].push((node instanceof XmlNode) ? node.el : node);
//  }
  return this;

}


XmlNode.prototype.setChild = function (tag, node) {
  this.el[tag] = (node instanceof XmlNode) ? node.el : node;
  return this;
}