"use strict";


module.exports = {
  Presentation: require('./presentation'),
  Slide: require('./slide'),
  Shape: require('./shape'),
  spPr: require('./spPr'),
  emu: {
    inch: function(val) { return val * 914400; },
    point: function(val) { return val * 914400 / 72; },
    px: function(val) { return val * 914400 / 72; },
    cm: function(val) { return val * 360000; }
  }
};




