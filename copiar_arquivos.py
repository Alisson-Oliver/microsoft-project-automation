from pathlib import Path
import shutil

from config import SUFIXO_COPIA, PASTA_COPIA

def copiar_mpp(origem: Path, sufixo: str = SUFIXO_COPIA) -> Path:
    pasta_copia = Path(PASTA_COPIA)
    pasta_copia.mkdir(parents=True, exist_ok=True)
    
    copia = pasta_copia / (origem.stem + sufixo + origem.suffix)
    
    shutil.copy2(origem, copia)
    print(f"[COPIA] {origem.name} -> {copia.name}")
    
    return copia

def copiar_lista(arquivos: list[Path]) -> list[Path]:
    copias = []
    for a in arquivos:
        copia = copiar_mpp(a)
        copias.append(copia)
    return copias

def deletar_copia(copias: list[Path]) -> int:
    deletados = 0
    for copia in copias:
        try:
            copia.unlink()
            print(f"[DELETADO] {copia.name}")
            deletados += 1
        except Exception as e:
            print(f"[ERRO] Não foi possível deletar {copia.name}: {e}")
    return deletados

if __name__ == "__main__": 
    from buscar_arquivos import listar_mpp
    
    encontrados = listar_mpp()
    copias = copiar_lista(encontrados)
    print(f"\n[OK] {len(copias)} cópia(s) criada(s).")