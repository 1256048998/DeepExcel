"""The runtime dependency list must actually cover what the app imports.

httpx was a runtime dependency of the model proxy but lived in
requirements-dev.txt, because pytest pulled it in locally. Nothing noticed until
the production image was built for the first time and the application failed to
import at startup. A test is cheaper than finding that out during a deploy.
"""

from __future__ import annotations

import ast
import pathlib
import re
import sys

APP = pathlib.Path(__file__).resolve().parents[1] / "app"
REQUIREMENTS = pathlib.Path(__file__).resolve().parents[1] / "requirements.txt"

# Import name -> distribution name, where they differ.
DISTRIBUTION_NAMES = {
    "jwt": "pyjwt",
    "sqlalchemy": "sqlalchemy",
    "multipart": "python-multipart",
    "email_validator": "email-validator",
}


def declared_distributions() -> set[str]:
    names = set()
    for line in REQUIREMENTS.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        names.add(re.split(r"[=<>\[]", line)[0].strip().lower())
    return names


def imported_top_level_modules() -> set[str]:
    modules = set()
    for path in APP.rglob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    modules.add(alias.name.split(".")[0])
            elif isinstance(node, ast.ImportFrom):
                # level > 0 is a relative import within the app itself.
                if node.level == 0 and node.module:
                    modules.add(node.module.split(".")[0])
    return modules


def test_every_third_party_import_is_declared():
    declared = declared_distributions()
    stdlib = set(sys.stdlib_module_names)

    missing = []
    for module in sorted(imported_top_level_modules()):
        if module in stdlib or module == "app":
            continue
        distribution = DISTRIBUTION_NAMES.get(module, module)
        if distribution.lower() not in declared:
            missing.append(f"{module} (would need '{distribution}')")

    assert not missing, (
        "app/ imports packages that requirements.txt does not declare, so the "
        "production image will fail to start:\n  " + "\n  ".join(missing)
    )
