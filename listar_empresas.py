"""Script auxiliar para listar empresas disponíveis na pasta de PDFs."""
from pathlib import Path

PASTA_RAIZ = r"C:\Users\weder\OneDrive\Área de Trabalho\codigos\iaindex\pdfs"

pasta = Path(PASTA_RAIZ)

if not pasta.exists():
    print(f"❌ Pasta não encontrada: {PASTA_RAIZ}")
    print("\nPor favor, crie a pasta e organize os PDFs na estrutura:")
    print("  pdfs/")
    print("    ├── Empresa1/")
    print("    │   ├── 2023/")
    print("    │   ├── 2024/")
    print("    │   └── 2025/")
    print("    └── Empresa2/")
    print("        └── ...")
else:
    empresas = [d.name for d in pasta.iterdir() if d.is_dir()]
    if empresas:
        print(f"✅ Empresas encontradas em {PASTA_RAIZ}:")
        for i, empresa in enumerate(empresas, 1):
            print(f"  {i}. {empresa}")
        print(f"\nPara processar apenas uma empresa, edite analisar_pdfs.py e defina:")
        print(f"  EMPRESA_FILTRO = \"{empresas[0]}\"")
    else:
        print(f"⚠️  Pasta existe mas não há subpastas (empresas) em: {PASTA_RAIZ}")
