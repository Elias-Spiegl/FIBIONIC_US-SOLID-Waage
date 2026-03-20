from __future__ import annotations

try:
    from .runtime_support import ensure_runtime_supported
except ImportError:
    from fibionic_scale_app.runtime_support import ensure_runtime_supported


def main() -> None:
    ensure_runtime_supported()
    try:
        from .app import main as app_main
    except ImportError:
        from fibionic_scale_app.app import main as app_main

    app_main()


__all__ = ["main"]


if __name__ == "__main__":
    main()
