"""Pseudonimização estável de titulares (CPF) e contas.

O mapa pseudônimo ↔ valor real é gravado só na pasta privada do gabarito (fora do repositório).
Ele é reaproveitado entre execuções para que o mesmo CPF/conta receba sempre o mesmo pseudônimo.
"""

import json
from pathlib import Path

from pydantic import BaseModel, Field

from coir.domain.cpf import so_digitos
from coir.erros import GabaritoError


class MapaPseudonimos(BaseModel):
    titulares: dict[str, str] = Field(default_factory=dict)  # cpf (só dígitos) -> TITULAR-n
    contas: dict[str, str] = Field(default_factory=dict)  # "CORRETORA|conta" -> CORRETORA-n

    @classmethod
    def carregar(cls, caminho: Path) -> "MapaPseudonimos":
        if not caminho.is_file():
            return cls()
        return cls.model_validate_json(caminho.read_text(encoding="utf-8"))

    def salvar(self, caminho: Path) -> None:
        caminho.parent.mkdir(parents=True, exist_ok=True)
        dados = self.model_dump()
        caminho.write_text(
            json.dumps(dados, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8"
        )

    def titular(self, cpf: str) -> str:
        digitos = so_digitos(cpf)
        if len(digitos) != 11:
            raise GabaritoError("CPF com quantidade de dígitos diferente de 11")
        if digitos not in self.titulares:
            self.titulares[digitos] = f"TITULAR-{len(self.titulares) + 1}"
        return self.titulares[digitos]

    def conta(self, corretora: str, conta: str) -> str:
        chave = f"{corretora}|{conta}"
        if chave not in self.contas:
            n = sum(1 for k in self.contas if k.startswith(f"{corretora}|")) + 1
            self.contas[chave] = f"{corretora}-{n}"
        return self.contas[chave]

    def __repr__(self) -> str:
        """Nunca exibe valores reais (o repr aparece em tracebacks e logs)."""
        return f"MapaPseudonimos({len(self.titulares)} titular(es), {len(self.contas)} conta(s))"

    __str__ = __repr__

    def valores_reais(self) -> list[str]:
        """CPFs e números de conta reais (para verificar que nenhum saiu em arquivo público)."""
        return [*self.titulares, *(k.split("|", 1)[1] for k in self.contas)]
