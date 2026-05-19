
import os
import glob
from pathlib import Path

from config import PASTA_RAIZ, BUSCA_RECURSIVA

def listar_mpp(pasta: str = PASTA_RAIZ, recursivo: bool = BUSCA_RECURSIVA) -> list[Path]: 
    padrao = "**/*.mpp" if recursivo else "*.mpp"
    arquivos = [
        Path(p) for p in glob.glob(os.path.join(pasta, padrao), recursive=recursivo)
    ]
    
    print(f"[BUSCA] {len(arquivos)} arquivo(s) .mpp encontrado(s) em: {pasta}")
    for a in arquivos:
        print(f"{a}")
    
    return arquivos

if __name__ == "__main__":
    listar_mpp()