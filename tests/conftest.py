"""Configuración compartida de los tests."""
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).resolve().parent / "fixtures"

if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))
