r"""Ponto de entrada do COIR.exe (janela, sem console): abre a interface na pasta de dados
escolhida (variável COIR_DADOS ou a configuração em %APPDATA%\COIR); sem nenhuma, o assistente
pede a pasta."""

import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()  # a janela nativa (pywebview) roda num processo à parte

    from coir.config import dados_dir
    from coir.db.sessao import NOME_BANCO
    from coir.servicos.estado import Estado
    from coir.ui.janela import iniciar

    iniciar(Estado(dados_dir() / NOME_BANCO))
