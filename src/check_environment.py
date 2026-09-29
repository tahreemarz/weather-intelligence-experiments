"""Quick environment check: run `python src/check_environment.py` after `pip install -e .[notebooks]`.

Verifies that the custom package (>= 0.2.0), the notebook helper `utstd` and the
Open-Meteo Historical Weather API are reachable before opening the notebooks.
"""

from __future__ import annotations

import sys


def main() -> int:
    ok = True
    try:
        import amla_package_26016360 as amla
        from amla_package_26016360 import weather  # noqa: F401

        major, minor = (int(x) for x in amla.__version__.split(".")[:2])
        assert (major, minor) >= (0, 2), amla.__version__
        print(f"[ok] amla-package-26016360 {amla.__version__} (weather subpackage available)")
    except Exception as exc:  # pragma: no cover - diagnostic script
        ok = False
        print(f"[!!] amla-package-26016360 >= 0.2.0 not importable: {exc!r}")
        print("     pip install --no-deps --index-url https://test.pypi.org/simple/ 'amla-package-26016360>=0.2.0'")

    for name in ("utstd", "xgboost", "sklearn", "pandas", "numpy", "matplotlib", "seaborn", "pyarrow", "nbformat"):
        try:
            mod = __import__(name)
            print(f"[ok] {name} {getattr(mod, '__version__', '')}")
        except ImportError:
            ok = False
            print(f"[!!] {name} missing -> pip install -e .[notebooks]")

    try:
        from amla_package_26016360.weather import SYDNEY, fetch_hourly

        df = fetch_hourly("2025-01-01", "2025-01-02", location=SYDNEY)
        print(f"[ok] Open-Meteo Historical Weather API reachable ({len(df)} hourly rows for 2025-01-01/02)")
    except Exception as exc:  # pragma: no cover - diagnostic script
        ok = False
        print(f"[!!] Open-Meteo not reachable: {exc!r}")

    print("environment OK" if ok else "environment has problems (see above)")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
