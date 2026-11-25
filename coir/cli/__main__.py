"""``python -m coir.cli``: a mesma linha de comando de ``uv run coir``."""

from coir.cli import main

if __name__ == "__main__":  # os processos de leitura (spawn) não rodam o comando de novo
    main()
