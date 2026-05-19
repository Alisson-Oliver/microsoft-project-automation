from buscar_arquivos import listar_mpp
from copiar_arquivos import copiar_lista, deletar_copia
from exportar_xlsx import iniciar_project, encerrar_project, exportar_lista


def main():
    print("=" * 60)
    print("  AUTOMAÇÃO MPP → XLSX")
    print("=" * 60)

    arquivos = listar_mpp()
    if not arquivos:
        print("\nNenhum arquivo .mpp encontrado. Encerrando.")
        return

    print()
    copias = copiar_lista(arquivos)

    print()
    app = iniciar_project()
    try:
        gerados = exportar_lista(app, copias)
    finally:
        encerrar_project(app)

    print()
    deletados = deletar_copia(copias)

    print()
    print("=" * 60)
    print(f"  Arquivos encontrados : {len(arquivos)}")
    print(f"  Cópias criadas       : {len(copias)}")
    print(f"  XLSX gerados         : {len(gerados)}")
    print(f"  Cópias deletadas     : {deletados}")
    print("=" * 60)


if __name__ == "__main__":
    main()