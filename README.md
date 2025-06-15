# Prerequisites

You will need the Python UV package manager:
- [uv](https://github.com/astral-sh/uv)

# Installation

- Clone the repository 
- `cd` into the cloned directory
- Run the following commands:

```bash
uv sync                         # to install the dependencies
source ./.venv/bin/activate     # to activate the virtual environment
```

# Running the code
```
uv run src/google_drive_v2.py
```

# Running tests
```
uv run pytest
```
