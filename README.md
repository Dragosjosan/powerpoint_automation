# PowerPoint Template Updater

Script to replace placeholders in PowerPoint templates with custom data (text, tables, images).

## Setup

```bash
# Create virtual environment
uv venv .venv --python 3.12
source .venv/bin/activate
# OR
.venv\Scripts\activate  # Windows

# Install dependencies
uv pip install -r pyproject.toml
```

## Usage
1. Create a PowerPoint [template](example-presentation.pptx)
   - Text placeholders: `{{variable_name}}`
   - Tables and image placeholders


2. Run script:
```bash
python main.py
```

## Input Data Format
```python
data = {
    "Slide Title": {
        "text": {"variable_name": "value"},
        "tables": {"0": {"data": [["row1"], ["row2"]]}},
        "images": {"0": "image_path.png"}
    }
}
```
