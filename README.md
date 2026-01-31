# Gerador de Jogos Mega-Sena (Tkinter) + Exportar PDF

Aplicativo em Python (Tkinter) para gerar jogos da Mega-Sena em 3 modos:
- Números mais sorteados (todo o período)
- Números mais sorteados (Mega da Virada - 31/12)
- Aleatórios

Também:
- Carrega resultados via planilha `.xlsx`
- Mostra preview de frequências
- Exporta jogos gerados para PDF

## Requisitos
- Python 3.10+ (recomendado)
- Dependências:
  - pandas
  - openpyxl
  - reportlab

## Instalação (Windows / Linux / macOS)
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

pip install -r requirements.txt

## Como usar

1. Coloque a planilha `Todos-Resultados-Mega-Sena...xlsx` no campo de seleção.
2. Escolha o modo de geração de jogos.
3. Clique em "Gerar Jogos".
4. Exporte para PDF quando terminar.
