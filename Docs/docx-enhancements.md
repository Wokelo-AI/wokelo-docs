# Enhanced python-docx Features Documentation

## Overview

This document covers the new enhanced features added to the python-docx package in the `wokelo_docs.docx` module. These enhancements extend the original functionality with additional capabilities for working with Word documents.

## Installation

```bash
pip install wokelo-docs
```

## Import

```python
from wokelo_docs.docx import Document
# Instead of: from docx import Document
```

## New Enhanced Features

### 1. Comments Support

The enhanced package adds comprehensive support for working with document comments.

#### Adding Comments to Runs

```python
from wokelo_docs.docx import Document
from datetime import datetime

# Create a new document
doc = Document()
paragraph = doc.add_paragraph("This is some text with a comment.")

# Get the first run
run = paragraph.runs[0]

# Add a comment to the run
comment = run.add_comment(
    text="This is a comment on the text",
    author="John Doe",
    initials="JD",
    dtime=datetime.now().isoformat()
)

doc.save('document_with_comments.docx')
```

#### Accessing Comments from Runs

```python
from wokelo_docs.docx import Document

# Open an existing document
doc = Document('document_with_comments.docx')

# Access comments from runs
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        comments = run.comments
        for comment in comments:
            print(f"Comment by {comment.author}: {comment.text}")
```

#### Document-level Comments Access

```python
# Access comments part directly
comments_part = doc.part._comments_part
# Work with comments at document level
```

### 2. Footnotes Support

Enhanced footnote functionality for better document structure.

#### Adding Footnotes

```python
from wokelo_docs.docx import Document

doc = Document()
paragraph = doc.add_paragraph("This text has a footnote.")

# Access footnotes part
footnotes_part = doc.part._footnotes_part
# Add footnote functionality (implementation details in footnotes module)
```

#### Accessing Footnotes from Runs

```python
# Check if a run has a footnote
for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        footnote_text = run.footnote
        if footnote_text:
            print(f"Footnote: {footnote_text}")
```

### 3. Enhanced Chart Support

Improved chart embedding capabilities with better integration.

#### Adding Charts to Documents

```python
from wokelo_docs.docx import Document
from wokelo_docs.docx.chart.data import CategoryChartData
from wokelo_docs.docx.enum.chart import XL_CHART_TYPE

# Create document and chart data
doc = Document()
chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('Sales', (100, 150, 120, 180))

# Add chart to document
paragraph = doc.add_paragraph()
run = paragraph.add_run()

# Add chart with specific dimensions and position
chart = run.add_chart(
    chart_type=XL_CHART_TYPE.COLUMN_CLUSTERED,
    x=0,
    y=0,
    cx=5000000,  # Width in EMUs
    cy=3000000,  # Height in EMUs
    chart_data=chart_data
)

doc.save('document_with_chart.docx')
```

#### Chart Integration in Runs

```python
# Charts are now properly integrated as inline shapes in runs
# Allowing for better text flow and positioning
```

### 4. Hyperlink Detection and Management

Enhanced hyperlink handling capabilities.

#### Detecting Hyperlinks in Runs

```python
from wokelo_docs.docx import Document

doc = Document('document_with_links.docx')

for paragraph in doc.paragraphs:
    for run in paragraph.runs:
        if run.is_hyperlink:
            link_target, is_external = run.get_hyperLink()
            if is_external:
                print(f"External link: {link_target}")
            else:
                print(f"Internal anchor: {link_target}")
```

#### Working with Hyperlink Properties

```python
# Check if run is part of a hyperlink
if run.is_hyperlink:
    # Get hyperlink details
    link_info = run.get_hyperLink()
    # Returns tuple: (link_target, is_external_boolean)
```


## Advanced Usage Examples

### Complete Comment Workflow

```python
from wokelo_docs.docx import Document
from datetime import datetime

# Create document with comprehensive comment usage
doc = Document()

# Add content
p1 = doc.add_paragraph("This is the introduction paragraph.")
p2 = doc.add_paragraph("This paragraph contains important information.")

# Add comments to specific runs
intro_run = p1.runs[0]
intro_comment = intro_run.add_comment(
    text="Consider expanding this introduction",
    author="Reviewer 1",
    initials="R1",
    dtime=datetime.now().isoformat()
)

important_run = p2.runs[0]
important_comment = important_run.add_comment(
    text="This needs fact-checking",
    author="Editor",
    initials="ED"
)

# Save and later read comments
doc.save('reviewed_document.docx')

# Read comments back
doc2 = Document('reviewed_document.docx')
for paragraph in doc2.paragraphs:
    for run in paragraph.runs:
        if run.comments:
            for comment in run.comments:
                print(f"{comment.author}: {comment.text}")
```

### Enhanced Chart and Media Integration

```python
from wokelo_docs.docx import Document
from wokelo_docs.docx.shared import Inches

doc = Document()

# Add title
title = doc.add_heading('Sales Report', 0)

# Add description
desc = doc.add_paragraph("Below is the quarterly sales chart:")

# Create and add chart
paragraph = doc.add_paragraph()
run = paragraph.add_run()

# Chart with proper dimensions
chart = run.add_chart(
    chart_type=XL_CHART_TYPE.LINE,
    x=0,
    y=0,
    cx=Inches(6),
    cy=Inches(4),
    chart_data=chart_data
)

# Add image alongside
img_para = doc.add_paragraph()
img_run = img_para.add_run()
img_run.add_picture('chart_image.png', width=Inches(3))

doc.save('comprehensive_report.docx')
```


## Support

For issues and feature requests, please visit the project repository or contact support@wokelo.ai