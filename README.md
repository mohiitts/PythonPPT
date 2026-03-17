# OOP Pillars PPT Generator

A Python script that automatically generates a complete, professionally formatted
17-slide PowerPoint presentation on the **4 Pillars of Object-Oriented Programming**
(Encapsulation, Abstraction, Inheritance, Polymorphism).

## What the script does

Running `generate_oops_ppt.py` creates `OOP_Pillars_Presentation.pptx` containing:

- A title slide with the date
- Concept slides explaining each of the 4 OOP pillars
- Java code examples rendered in **Courier New** on a light-gray background
- Comparison tables with a blue header row and alternating row colours
- A memory-trick summary slide ("A PIE")
- Slide numbers on every slide (except the title)

## Prerequisites

- Python 3.x

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python generate_oops_ppt.py
```

## Output

The script generates **`OOP_Pillars_Presentation.pptx`** in the current directory and
prints a confirmation message:

```
✅ Presentation generated successfully: OOP_Pillars_Presentation.pptx
```