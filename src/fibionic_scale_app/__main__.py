from __future__ import annotations

try:
    from .app import main
except ImportError:
    from fibionic_scale_app.app import main

__all__ = ["main"]


if __name__ == "__main__":
    main()
