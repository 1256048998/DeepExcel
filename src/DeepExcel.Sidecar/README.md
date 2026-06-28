# DeepExcel Sidecar

Python sidecar process for DeepExcel. Runs Claude Agent SDK and communicates with the C# COM add-in via JSON Lines over stdin/stdout.

## Development

```bash
cd src/DeepExcel.Sidecar
pip install -r requirements.txt
pytest tests/ -v
```

## Run standalone (debug)

```bash
python sidecar.py
# Reads JSON Lines from stdin, writes JSON Lines to stdout
```
