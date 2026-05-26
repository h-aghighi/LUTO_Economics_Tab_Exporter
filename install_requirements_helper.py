#!/usr/bin/env python3
"""
install_requirements_helper.py

Helper script for installing the Python packages required by:

LUTO Economics Exporters

Required packages:
- pandas
- openpyxl

This script tries several installation methods:
1. python -m pip install pandas openpyxl
2. python -m pip install --user pandas openpyxl
3. py -m pip install pandas openpyxl
4. py -m pip install --user pandas openpyxl
5. conda install -y pandas openpyxl
6. conda install -y -c conda-forge pandas openpyxl

Run from PowerShell:

python .\install_requirements_helper.py
"""

import importlib.util
import shutil
import subprocess
import sys
from typing import List, Tuple


REQUIRED_PACKAGES: List[Tuple[str, str]] = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
]


def module_exists(module_name: str) -> bool:
    """Return True if a Python module can be imported."""
    return importlib.util.find_spec(module_name) is not None


def missing_packages() -> List[str]:
    """Return package names whose import modules are missing."""
    missing = []

    for package_name, import_name in REQUIRED_PACKAGES:
        if not module_exists(import_name):
            missing.append(package_name)

    return missing


def run_command(command: List[str]) -> bool:
    """Run a command and return True if it succeeds."""
    print("\nTrying command:")
    print(" ".join(command))

    try:
        result = subprocess.run(command, check=False)
        return result.returncode == 0
    except FileNotFoundError:
        print(f"Command not found: {command[0]}")
        return False
    except Exception as exc:
        print(f"Command failed unexpectedly: {exc}")
        return False


def install_required_packages() -> bool:
    """
    Try multiple installation methods for missing packages.

    Returns:
        True if all packages are available after installation attempts.
        False otherwise.
    """
    missing = missing_packages()

    if not missing:
        print("All required packages are already installed.")
        return True

    print("Missing packages:")
    print(", ".join(missing))

    strategies: List[List[str]] = []

    # Strategy 1: current Python pip
    strategies.append([sys.executable, "-m", "pip", "install", *missing])

    # Strategy 2: current Python pip with --user
    strategies.append([sys.executable, "-m", "pip", "install", "--user", *missing])

    # Strategy 3 and 4: Windows Python launcher, if available
    if shutil.which("py"):
        strategies.append(["py", "-m", "pip", "install", *missing])
        strategies.append(["py", "-m", "pip", "install", "--user", *missing])

    # Strategy 5 and 6: Conda, if available
    if shutil.which("conda"):
        strategies.append(["conda", "install", "-y", *missing])
        strategies.append(["conda", "install", "-y", "-c", "conda-forge", *missing])

    for command in strategies:
        success = run_command(command)

        if success:
            still_missing = missing_packages()

            if not still_missing:
                print("\nPackages installed successfully.")
                return True

            print("\nCommand completed, but these packages are still missing:")
            print(", ".join(still_missing))

    print("\nAutomatic installation failed.")
    print("\nTry one of these manually in PowerShell:")
    print("python -m pip install pandas openpyxl")
    print("python -m pip install --user pandas openpyxl")
    print("py -m pip install pandas openpyxl")
    print("py -m pip install --user pandas openpyxl")
    print("conda install pandas openpyxl")
    print("conda install -c conda-forge pandas openpyxl")

    return False


def verify_required_packages() -> bool:
    """Verify imports after installation."""
    print("\nVerifying required packages...")

    ok = True

    for package_name, import_name in REQUIRED_PACKAGES:
        if module_exists(import_name):
            print(f"OK: {package_name}")
        else:
            print(f"MISSING: {package_name}")
            ok = False

    return ok


def main() -> None:
    print("LUTO Economics Exporters - Requirements Helper")
    print("Python executable:")
    print(sys.executable)

    installed = install_required_packages()
    verified = verify_required_packages()

    if installed and verified:
        print("\nReady. You can now run the exporter.")
        sys.exit(0)

    print("\nNot ready. Required packages are still missing.")
    sys.exit(1)


if __name__ == "__main__":
    main()