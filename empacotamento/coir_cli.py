"""Ponto de entrada do coir-cli.exe (console): a mesma linha de comando de `uv run coir`."""

import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()

    from coir.cli import main

    main()
