# Developer environment

Set up a local Python virtual environment to isolate dependencies.

## Create venv (once)

```powershell
python -m venv .venv
```

## Activate venv

- Windows PowerShell
```powershell
./.venv/Scripts/Activate.ps1
```

- Windows cmd
```cmd
.venv\Scripts\activate.bat
```

- Linux/macOS bash/zsh
```bash
source .venv/bin/activate
```

## Deactivate venv

```bash
deactivate
```

After activation, install dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```