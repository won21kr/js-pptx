# js-pptx API


# Classes
* Presentation
* Slide
* ShapeTree
* Shape
* SlideSize
* Chart
* SlideMaster
* SlideLayout
* ShapeProperties
* NonVisualShapeProperties
* Properties
* NonVisualProperties
* Background
* BackgroundProperties
* BackgroundReference

* ColorScheme
* ThemeElements
* ColorMapOverride

* TextBody
* Paragraph
* HeaderFooter


* Theme
* Slides



# Design questions


* How to keep track of relationships?
   - Does the Chart class register its rels and parts in the Presentation object?
   - Or does the Presentation keep track of the rels and provide a registration method?
   - Who owns the master object?
*