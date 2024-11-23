# Extract PPT Slides
Extract all slides, text and images from PPT-files

# Usage with uv
See pyproject.toml for info or install uv
```
uv sync

uv run extract_ppt.py
```

Thank you to developers of various libraries that makes and claude.ai 

dependencies = [
    "comtypes>=1.4.8",
    "pillow>=11.0.0",
    "python-pptx>=1.0.2",
    "pywin32>=308",
]
