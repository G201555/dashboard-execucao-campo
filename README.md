# Dashboard de execução de campo

Pipeline completo de um painel de execução de campo: ruptura, mobiliário, roteiro, treinamentos, demonstração, vendas e giro, com estoque semanal das redes varejistas.

## Pipeline

| Etapa | Arquivo |
|---|---|
| ETL dos dados de campo no SQL Server | `tools/atualizar_dashboard.ps1` |
| Importação dos arquivos semanais de estoque e venda das redes, com de-para de produtos | `tools/importar_estoque_semanal.ps1`, `tools/importar_estoque_varejo.py` |
| Cruzamento da base de lojas das redes com o cadastro | `tools/cruzar_base_lojas.py` |
| Publicação no Firestore | `tools/publicar_firestore.js` |
| Tudo num comando: ETL, importação, Firestore e deploy | `tools/atualizar_e_publicar.ps1` |

O acesso aos dados é controlado por login (Firebase Auth) e por uma lista de e-mails nas regras do Firestore.

## O que vale olhar

- Arquivos de varejo chegam cada semana num formato diferente por rede; o importador normaliza tudo para um modelo só.
- O histórico semanal é acumulado e nunca sobrescrito.

## Stack

PowerShell, Python, openpyxl, Node.js, SQL Server, Firebase Auth, Firestore, Hosting.

---
Projeto feito no trabalho, publicado sem dados do cliente. A página principal do painel ficou de fora porque é gerada com os dados embutidos.
