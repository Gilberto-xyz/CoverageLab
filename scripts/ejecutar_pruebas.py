"""Ejecuta las pruebas del estudio desde cualquier carpeta de trabajo."""
from pathlib import Path
import os
import sys
import unittest

root = Path(__file__).resolve().parent
os.chdir(root)
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
sys.dont_write_bytecode = True
temporary = root / 'outputs'
existed = temporary.exists()
temporary.mkdir(exist_ok=True)
try:
    unittest.main(module=None, argv=[sys.argv[0], 'discover', '-s', 'tests', *sys.argv[1:]])
finally:
    if not existed and temporary.is_dir() and not any(temporary.iterdir()):
        temporary.rmdir()
