# Enhanced python-pptx Features Documentation

## Overview

This document covers the new enhanced features added to the python-pptx package in the `wokelo_docs.pptx` module. These enhancements extend the original functionality with additional capabilities for working with PowerPoint presentations, particularly focusing on slide masters, advanced shape management, and improved performance features.

## Installation

```bash
pip install wokelo-docs
```

## Import

```python
from wokelo_docs.pptx import Presentation
# Instead of: from pptx import Presentation
```

## New Enhanced Features

### 1. Enhanced Slide Master Support

The enhanced package adds comprehensive support for working with slide masters, including the ability to add shapes directly to slide masters.

#### Adding Pictures to Slide Masters

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.util import Inches

# Create or open a presentation
prs = Presentation()
slide_master = prs.slide_master

# Add a picture to the slide master
# This will appear on all slides using this master
picture = slide_master.shapes.add_picture(
    'company_logo.png',
    left=Inches(0.5),
    top=Inches(0.5),
    width=Inches(1),
    height=Inches(0.5)
)

# The picture will now appear on all slides using this master
slide = prs.slides.add_slide(prs.slide_layouts[0])
prs.save('presentation_with_master_logo.pptx')
```

#### Adding Shapes to Slide Masters

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.enum.shapes import MSO_SHAPE
from wokelo_docs.pptx.util import Inches

prs = Presentation()
slide_master = prs.slide_master

# Add a rectangle shape to the master
shape = slide_master.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    left=Inches(8),
    top=Inches(6),
    width=Inches(1.5),
    height=Inches(0.5)
)

# Customize the shape
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0, 51, 102)  # Corporate blue
shape.text_frame.text = "Confidential"

prs.save('presentation_with_master_watermark.pptx')
```

#### Adding Connectors to Slide Masters

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.enum.shapes import MSO_CONNECTOR_TYPE
from wokelo_docs.pptx.util import Inches

prs = Presentation()
slide_master = prs.slide_master

# Add a connector line to the master
connector = slide_master.shapes.add_connector(
    MSO_CONNECTOR_TYPE.STRAIGHT,
    begin_x=Inches(0),
    begin_y=Inches(7),
    end_x=Inches(10),
    end_y=Inches(7)
)

# Customize the connector
connector.line.color.rgb = RGBColor(128, 128, 128)
connector.line.width = Pt(2)

prs.save('presentation_with_master_line.pptx')
```

### 2. Turbo-Add Performance Mode

Enhanced performance feature for adding large numbers of shapes to slides with significant speed improvements.

#### Enabling Turbo-Add Mode

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.enum.shapes import MSO_SHAPE
from wokelo_docs.pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# Enable turbo-add mode for performance
slide.shapes.turbo_add_enabled = True

# Now add many shapes efficiently
for i in range(100):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(i * 0.1),
        top=Inches(1),
        width=Inches(0.5),
        height=Inches(0.5)
    )

# Disable turbo mode when done
slide.shapes.turbo_add_enabled = False

prs.save('presentation_with_many_shapes.pptx')
```

#### Performance Comparison

```python
import time
from wokelo_docs.pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# Without turbo mode (slower)
start_time = time.time()
for i in range(50):
    slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(i*0.1), Inches(1), Inches(0.5), Inches(0.5))
normal_time = time.time() - start_time

# Clear slide and test with turbo mode
slide.shapes.turbo_add_enabled = True
start_time = time.time()
for i in range(50):
    slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(i*0.1), Inches(2), Inches(0.5), Inches(0.5))
turbo_time = time.time() - start_time

print(f"Normal mode: {normal_time:.2f}s")
print(f"Turbo mode: {turbo_time:.2f}s")
print(f"Speed improvement: {normal_time/turbo_time:.1f}x faster")
```

### 3. Enhanced Shape ID Management

Improved shape ID generation and management for better performance and reliability.

#### Smart Shape ID Generation

```python
from wokelo_docs.pptx import Presentation

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# The enhanced version automatically manages shape IDs more efficiently
# Get the next available shape ID
next_id = slide.shapes._next_shape_id
print(f"Next shape will have ID: {next_id}")

# Maximum shape ID tracking
max_id = slide.shapes._spTree.max_shape_id
print(f"Current maximum shape ID: {max_id}")

# Add a shape and see the ID management in action
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
print(f"Added shape with ID: {shape.shape_id}")
```

### 4. Advanced Group Shape Management

Enhanced group shape functionality with improved extent calculation and positioning.

#### Creating and Managing Group Shapes

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.enum.shapes import MSO_SHAPE
from wokelo_docs.pptx.util import Inches

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# Create individual shapes first
shape1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(1), Inches(1))
shape2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2.5), Inches(1), Inches(1), Inches(1))
shape3 = slide.shapes.add_shape(MSO_SHAPE.TRIANGLE, Inches(1.75), Inches(2.5), Inches(1), Inches(1))

# Create a group with the shapes
group = slide.shapes.add_group_shape([shape1, shape2, shape3])

# The group automatically calculates its extents based on contained shapes
print(f"Group position: ({group.left}, {group.top})")
print(f"Group size: {group.width} x {group.height}")

# Add more shapes to the group
group_shapes = group.shapes
new_shape = group_shapes.add_shape(
    MSO_SHAPE.STAR_5_POINT,
    Inches(0.5),  # Relative to group
    Inches(0.5),
    Inches(0.5),
    Inches(0.5)
)

prs.save('presentation_with_groups.pptx')
```


## Advanced Usage Examples

### Complete Slide Master Customization

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.util import Inches, Pt
from wokelo_docs.pptx.dml.color import RGBColor
from wokelo_docs.pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR_TYPE

# Create presentation with comprehensive master customization
prs = Presentation()
slide_master = prs.slide_master

# Add corporate logo to master
logo = slide_master.shapes.add_picture(
    'company_logo.png',
    left=Inches(0.2),
    top=Inches(0.2),
    width=Inches(1.2),
    height=Inches(0.6)
)

# Add footer line
footer_line = slide_master.shapes.add_connector(
    MSO_CONNECTOR_TYPE.STRAIGHT,
    begin_x=Inches(0),
    begin_y=Inches(7),
    end_x=Inches(10),
    end_y=Inches(7)
)
footer_line.line.color.rgb = RGBColor(50, 50, 50)
footer_line.line.width = Pt(1)

# Add watermark
watermark = slide_master.shapes.add_shape(
    MSO_SHAPE.RECTANGLE,
    left=Inches(8),
    top=Inches(6.5),
    width=Inches(1.5),
    height=Inches(0.4)
)
watermark.fill.solid()
watermark.fill.fore_color.rgb = RGBColor(240, 240, 240)
watermark.line.fill.background()
watermark.text_frame.text = "CONFIDENTIAL"

# Create slides that inherit these elements
slide1 = prs.slides.add_slide(prs.slide_layouts[0])
slide1.shapes.title.text = "Enhanced Features Demo"

slide2 = prs.slides.add_slide(prs.slide_layouts[1])
slide2.shapes.title.text = "All Slides Inherit Master Elements"

prs.save('corporate_presentation.pptx')
```

### High-Performance Shape Generation

```python
from wokelo_docs.pptx import Presentation
from wokelo_docs.pptx.util import Inches
from wokelo_docs.pptx.enum.shapes import MSO_SHAPE
import random

prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

# Enable turbo mode for maximum performance
slide.shapes.turbo_add_enabled = True

# Generate a complex visualization with many shapes
colors = [RGBColor(255, 0, 0), RGBColor(0, 255, 0), RGBColor(0, 0, 255), 
          RGBColor(255, 255, 0), RGBColor(255, 0, 255), RGBColor(0, 255, 255)]

shapes_data = []
for i in range(200):
    x = random.uniform(0.5, 9)
    y = random.uniform(0.5, 6.5)
    size = random.uniform(0.1, 0.3)
    
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        left=Inches(x),
        top=Inches(y),
        width=Inches(size),
        height=Inches(size)
    )
    
    # Customize appearance
    color = random.choice(colors)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    
    shapes_data.append((x, y, size, color))

# Connect some shapes with lines
slide.shapes.turbo_add_enabled = True  # Keep turbo mode for connectors too
for i in range(0, len(shapes_data)-1, 10):
    x1, y1, _, _ = shapes_data[i]
    x2, y2, _, _ = shapes_data[i+1]
    
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR_TYPE.STRAIGHT,
        begin_x=Inches(x1),
        begin_y=Inches(y1),
        end_x=Inches(x2),
        end_y=Inches(y2)
    )
    connector.line.color.rgb = RGBColor(128, 128, 128)

slide.shapes.turbo_add_enabled = False
prs.save('complex_visualization.pptx')
```


- Compatible with existing python-pptx code
- Requires Python 3.6+
- Works with PowerPoint 2010+ presentation formats
- Maintains backward compatibility with existing .pptx files
- Enhanced features gracefully degrade in older PowerPoint versions

## Support

For issues and feature requests, please visit the project repository or contact support@wokelo.ai