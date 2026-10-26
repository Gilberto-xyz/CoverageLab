"""Separación entre código y modelos en scripts/ y datos en recursos/."""
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parent


def data_directory():
    if SCRIPTS.name == 'scripts' and (SCRIPTS.parent / 'README.md').is_file():
        return SCRIPTS.parent / 'recursos'
    return SCRIPTS


def coverage_template(root):
    local = Path(root).resolve() / 'Modelo_PPT.pptx'
    return local if local.is_file() else SCRIPTS / 'Modelo_PPT.pptx'


def external_research_directory(root):
    root = Path(root).resolve()
    project = root.parent if root.name == 'recursos' and (root.parent / 'README.md').is_file() else root
    return project.parent / '5W1H'


def state_directory(root):
    return Path(root).resolve() / '.estado'
