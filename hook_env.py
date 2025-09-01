# hook_env.py
import os, sys
from pathlib import Path

def _load_env():
    try:
        from dotenv import load_dotenv
    except Exception:
        load_dotenv = None

    candidates = []
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).parent
        candidates.append(exe_dir / ".env")                # .env рядом с exe (приоритет)
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            candidates.append(Path(meipass) / ".env")      # .env внутри бандла
    candidates.append(Path.cwd() / ".env")                 # на всякий

    for p in candidates:
        if p.exists():
            if load_dotenv:
                load_dotenv(p, override=False)             # не затираем уже выставленные переменные
            else:
                # простой fallback-парсер
                for line in p.read_text(encoding="utf-8").splitlines():
                    s = line.strip()
                    if not s or s.startswith("#") or "=" not in s:
                        continue
                    k, v = s.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())
            break

_load_env()
