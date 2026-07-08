"""Cria cópias de homologação das calculadoras de PU MtM.

Para cada ativo com CalcPath (exceto liquidados), copia o arquivo Excel para o
mesmo diretório adicionando " - Homologação" ao nome, antes da extensão.

Exemplo:
    Ativo DI+Spread A s1.xlsx  →  Ativo DI+Spread A s1 - Homologação.xlsx

Uso:
    python -m mtm_skills.criar_homologacao

Requisitos:
    - cadastro_ativos.xlsx fechado no Excel antes de rodar
"""
import glob
import hashlib
import shutil
import sys
from pathlib import Path

_RAIZ = Path(__file__).resolve().parents[1]
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

from pu_mtm.app import config
from pu_mtm.dados.csvio import ler_dict


def _glob1(pattern: str) -> str | None:
    hits = glob.glob(pattern, recursive=True)
    return hits[0] if hits else None


CADASTRO = str(config.CADASTRO)
FLUXO_DIR = config.FLUXO_DIR
SUFIXO = " - Homologação"


def _md5(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def criar_homologacao(
    cadastro_path: str | None = None,
    fluxo_dir: Path | None = None,
    dry_run: bool = False,
) -> None:
    cadastro_path = cadastro_path or CADASTRO
    fluxo_dir = fluxo_dir or FLUXO_DIR

    try:
        rows = ler_dict(cadastro_path)
    except PermissionError:
        raise SystemExit(
            f"\nArquivo aberto no Excel — feche-o antes de rodar:\n  {cadastro_path}"
        )

    FAMILIAS_CDI = {"di_puro", "di_spread"}
    ativos = [
        r for r in rows
        if r.get("CalcPath")
        and r.get("Status", "").lower() != "liquidado"
        and r.get("Familia", "").lower() in FAMILIAS_CDI
    ]

    print(f"Cadastro : {cadastro_path}")
    print(f"FluxoDir : {fluxo_dir}")
    print(f"Ativos   : {len(ativos)} (com CalcPath, excluindo liquidados)")
    if dry_run:
        print("(DRY RUN — nenhum arquivo será criado)\n")
    else:
        print()

    concluidas = com_diferencas = atualizadas = faltam = nao_achado = erros = 0

    for r in ativos:
        origem = Path(fluxo_dir) / r["CalcPath"]
        destino = origem.with_stem(origem.stem + SUFIXO)
        apelido = r.get("Apelido", r["IdSerie"])

        if not origem.exists():
            print(f"  [{r['IdSerie']:>6}] NÃO ENCONTRADO : {origem.name}")
            nao_achado += 1
            continue

        ja_existe = destino.exists()
        diferente = ja_existe and (_md5(origem) != _md5(destino))

        if dry_run:
            if not ja_existe:
                print(f"  [{r['IdSerie']:>6}] FALTA          : {destino.name}")
                faltam += 1
            elif diferente:
                print(f"  [{r['IdSerie']:>6}] COM DIFERENÇAS : {destino.name}")
                com_diferencas += 1
            else:
                print(f"  [{r['IdSerie']:>6}] CONCLUÍDA      : {destino.name}")
                concluidas += 1
            continue

        try:
            shutil.copy2(str(origem), str(destino))
            if ja_existe:
                print(f"  [{r['IdSerie']:>6}] ATUALIZADA     : {apelido}")
                atualizadas += 1
            else:
                print(f"  [{r['IdSerie']:>6}] CRIADA         : {apelido}")
                concluidas += 1
        except Exception as e:
            print(f"  [{r['IdSerie']:>6}] ERRO           : {apelido} — {e}")
            erros += 1

    print()
    if dry_run:
        print(f"  Concluídas        : {concluidas}")
        print(f"  Existem diferenças: {com_diferencas}")
        print(f"  Faltam            : {faltam}")
        print(f"  Não achadas       : {nao_achado}")
    else:
        print(f"  Criadas    : {concluidas}")
        print(f"  Atualizadas: {atualizadas}")
        print(f"  Não achadas: {nao_achado}")
        print(f"  Erros      : {erros}")


if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    criar_homologacao(dry_run=dry_run)
