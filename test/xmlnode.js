var XmlNode = require('../lib/xmlnode')
var assert = require('assert');

describe("XmlNode", function () {
  it('constructs with new', function () {
    var node = new XmlNode("p:sld", {})
    assert(node.tag == 'p:sld')
  });

  it('constructs as function', function () {
    var node = XmlNode("p:sld", {})
    assert(node.tag == 'p:sld')
  });

  it('constructs as function', function () {
    var node = new XmlNode("p:sld")
    node.attr('color', 'red')
    assert(node.tag == 'p:sld');
    assert(node.attr('color') == 'red')
  });

  it('constructs as function', function () {
    var node =  XmlNode("p:sld")
    node.attr('color', 'red')
    assert(node.tag == 'p:sld');
    assert(node.attr('color') == 'red')
  });


  it('constructs as function', function () {
    var node = XmlNode("p:sld").attr('color', 'red').attr({ size: 3, transparent: false})
    assert(node.tag == 'p:sld');
    assert(node.attr('color') == 'red')
    assert(node.attr('size') == 3)
    assert(node.attr('transparent') == false)
  })

})