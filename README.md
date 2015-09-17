# js-pptx
Pure Javascript reader/writer/editor for PowerPoint, for use in Node.js or the browser.


# Design goals
* Read/edit/author PowerPoint .pptx files
* Pure Javascript with clean IP
* Run in browser and/or Node.js
* Friendly API for basic tasks, like text, shapes, charts, tables
* Access to raw XML for when you need to be very specific
* Rigorous testing


# Current status
Early in development.  It can currently:
 * read an existing PPTX file
 * retain all existing content
 * add slides, shapes, and charts
 * save as a PPTX file
 * basic unit tests

What it cannot yet do is:
 * Programmatically retrieve / query / edit existing slides
 * Generate themes, layouts, masters, animations, etc.

# License
GNU General Public License (GPL)

# Install

In node.js
```
npm install protobi/js-pptx
```

In the browser: **(Not yet implemented)**
```
<script src="js-pptx.js"></script>
```

# Dependencies
* [xml2js](https://github.com/nfarina/xmldoc)
* [async](https://github.com/caolan/async)
* [jszip](https://stuk.github.io/jszip)

# Usage

```js
var PPTX = require('../lib/pptx');
var fs = require('fs');

var INFILE = './test/files/parts3.pptx';
var OUTFILE = './test/files/parts3-a.pptx';

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;

  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {

    var slide1 = pptx.getSlide('slide1');
    var shapes = slide1.getShapes();

    var shapes = slide1.getShapes()
    shapes[3]
        .text("Now it's a trapezoid")
        .shapeProperties()
        .x(PPTX.emu.inch(1))
        .y(PPTX.emu.inch(1))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(0.75))
        .prstGeom('trapezoid');
    });
  });
});

```

# Inspiration / Motivation
Inspired by [officegen](https://github.com/ZivBarber/officegen),
which creates pptx with text/shapes/images/tables/charts wonderfully (but does not read existing PPT files).

Also inspired by [js-xlsx](https://github.com/SheetJS/js-xlsx)
which reads/writes XLSX/XLS/XLSB, works in the browser and Node.js, and has an incredibly
thorough test suite (but does not read or write PowerPoint).

Motivated by desire to read and modify existing presentations, to inherit their themes, layouts and possibly content,
and work in the browser if possible.

https://github.com/protobi/js-pptx/wiki/API

# Design Philosophy
The design concept is to represent the Office document at two levels of abstraction:
* **Raw XML**  The actual complete OpenXML representation, in all its detail
* **Conceptual classes**  Simple Javascript classes that provide a convenient API

The conceptual classes provides a clear simple way to do common tasks, e.g. `Presentation().addSlide().addChart(data)`.

The raw API provides a way to do anything that the OpenXML allows, even if it's not yet in the conceptual classes, e.g.
e.g. `Presentation.getSlide(3).getShape(4).get('a:prstGeom').attr('prst', 'trapezoid')`


This solves a major dilemma in existing projects, which have many issue reports like "Please add this crucial feature to the API".
By being able to access the raw XML, all the features in OpenXML are available, while we make many of them more convenient.

The technical approach here uses:
* `JSZip` to unzip an existing `.pptx` file and zip it back,
* `xml2js` to convert the XML to Javascript and back to XML.

Converting to Javascript allows the content to be manipulated programmatically.  For each major entity, a Javascript class is created,
such as:
 * PPTX.Presentation
 * PPTX.Slide
 * PPTX.Shape
 * PPTX.spPr  // ShapeProperties
 * etc.

These classes allow properties to be set, and chained in a manner similar to d3 or jQuery.
The Javascript classes provide syntactic sugar, as a convenient way to query and modify the presentation.

But we can't possibly create a Javascript class that covers every entity and option defined in OpenXML.
So each of these classes exposes the  XML-to-Javascript object as a property `.content`, giving you theoretically
direct access to anything in the OpenXML standard, enabling you to take over
whenever the pre-defined features don't yet cover your particular use case.

It's up to you of course, to make sure that those changes convert to valid XML.  Debugging PPTX is a pain.

Right now, this uses English names for high-level constructs (e.g. `Presentation` and `Slide`),
but for lower level constructs uses names that directly mirror the OpenXML tagNames  (e.g.  `spPr` for ShapeProperties).

The challenge is it'll be a lot easier to extend the library if we follow the OpenXML tag names,
but the OpenXML tag names are so cryptic that they don't make great names for a Javascript library.

So we default to using the English name is used when returning objects even if the object has a cryptic class name, e.g.:
* `Slide.getShapes()` returns an array of `Shape` objects and
* `Shape.shapeProperties()` returns an `spPr` object.

Ideally would be consistent, and am working out which way to go.  Advice is welcome!

This library currently assumes it's starting from an existing presentation, and doesn't (yet) create one from scratch.
This allows you to use existing themes, styles and layouts.


# License
GNU General Public License (GPL)

# Install

In node.js
```
npm install protobi/js-pptx
```

In the browser:
```
<script src="js-pptx.js"></script>  // will use browserify but right now not yet implemented
```

# Dependencies
* [xml2js](https://github.com/nfarina/xmldoc)
* [async](https://github.com/caolan/async)
* [jszip](https://stuk.github.io/jszip)

# Usage

```js
"use strict";

var fs = require("fs");
var PPTX = require('..');


var INFILE = './test/files/minimal.pptx'; // a blank PPTX file with my layouts, themes, masters.
var OUTFILE = '/tmp/example.pptx';

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;
  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {
    var slide1 = pptx.getSlide('slide1');

    var slide2 = pptx.addSlide("slideLayout3"); // section divider
    var slide3 = pptx.addSlide("slideLayout6"); // title only


    var triangle = slide1.addShape()
        .text("Triangle")
        .shapeProperties()
        .x(PPTX.emu.inch(2))
        .y(PPTX.emu.inch(2))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(2))
        .prstGeom('triangle');

    var triangle = slide1.addShape()
        .text("Ellipse")
        .shapeProperties()
        .x(PPTX.emu.inch(4))
        .y(PPTX.emu.inch(4))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(1))
        .prstGeom('ellipse');

    for (var i = 0; i < 20; i++) {
      slide2.addShape()
          .text("" + i)
          .shapeProperties()
          .x(PPTX.emu.inch((Math.random() * 10)))
          .y(PPTX.emu.inch((Math.random() * 6)))
          .cx(PPTX.emu.inch(1))
          .cy(PPTX.emu.inch(1))
          .prstGeom('ellipse');
    }

    slide1.getShapes()[3]
        .text("Now it's a trapezoid")
        .shapeProperties()
        .x(PPTX.emu.inch(1))
        .y(PPTX.emu.inch(1))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(0.75))
        .prstGeom('trapezoid');

    var chart = slide3.addChart(barChart, function (err, chart) {

      fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
        if (err) throw err;
        console.log("open " + OUTFILE)
      });
    });
  });
});

var barChart = {
  title: 'Sample bar chart',
  renderType: 'bar',
  data: [
    {
      name: 'Series 1',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [4.3, 2.5, 3.5, 4.5]
    },
    {
      name: 'Series 2',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.4, 4.4, 1.8, 2.8]
    },
    {
      name: 'Series 3',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.0, 2.0, 3.0, 5.0]
    }
  ]
}
```

# Next steps
* Browserify and test in browser
* Publish to bower
* Add tables
* Add images
* Set presentation properties
* Set theme
* Set layouts
* Set masters

# Contribute

###Test:
`npm test`

###Build:
`npm run build`

###Minify:
`npm run minify`

###All:
`npm run all`

