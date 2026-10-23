# Metadata Review

Aplicação **Streamlit** para revisão de metadados YAML do Databricks Unity Catalog por Data Owners. Importa lotes YAML/YML/ZIP, organiza os objetos, valida tags, exporta Excel, reimporta propostas e gera novos YAMLs após confirmação. Processa arquivos na memória do computador que executa o servidor.

## Executar no PyCharm

O projeto existente `metadata-reviewer` foi utilizado, preservando suas configurações e o interpretador Python 3.12. Requer Python 3.11 ou superior.

1. Abra esta pasta no **PyCharm**.
2. Selecione `.venv/Scripts/python.exe` em **Settings → Project → Python Interpreter**.
3. Instale `requirements.txt` pelo gerenciador do interpretador. `requirements-lock.txt` reproduz as versões exatas testadas, incluindo ferramentas de teste.
4. Selecione a configuração compartilhada **Metadata Review** e clique em **Run**. Também é possível executar `run.py` pelo menu de contexto.
5. Abra **http://127.0.0.1:8501**.

Alternativamente, no terminal do projeto:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py
```

Também funciona com `python -m streamlit run streamlit_app.py` na pasta do projeto. A configuração restringe o servidor a `127.0.0.1`, desativa telemetria e observação automática de arquivos. Não há autenticação, banco de dados, serviços externos ou conexão com Databricks. O responsável preenchido é autodeclarado.

## Fluxo de revisão

1. Importe múltiplos arquivos YAML/YML ou ZIPs. Escolha adicionar ou substituir o lote. A aplicação informa arquivos inválidos, duplicidades, versões conflitantes e estruturas não reconhecidas. A exportação do lote inteiro fica bloqueada se houver erros de importação.
2. Navegue por catálogo, schema, tabela/view e coluna. Use pesquisa e filtros. A hierarquia vem dos arquivos; caminhos não criam tags ou objetos fictícios.
3. Revise descrições, nomes de negócio existentes e tags. Identificação técnica é somente leitura. Campos ausentes não recebem valores fictícios. Registre dúvidas e justificativas.
4. Quando disponíveis, carregue listas oficiais em JSON. As chaves são `nm_dom`, `nm_subdom`, `nm_sist_orig` e `ct_area_neg`, cada uma com uma lista de textos. As listas nunca são inferidas dos YAMLs. Sem referência, aparece **Lista oficial não carregada**.
5. Marque **Validation completed** individualmente. Os status são `Pending`, `In review`, `Awaiting clarification`, `Validated`, `Validated with remarks` e `Rejected`. Concluir uma tabela não conclui colunas nem outros objetos. Inconsistências técnicas e pendências permanecem visíveis.
6. Use **Export Excel for Data Owner**. As cinco abas visíveis são **Read Me**, **Objects**, **Tags**, **Columns** e **Approval**. O guia está em português. Identidades e valores atuais são protegidos; campos editáveis têm destaque. Há filtros, cabeçalhos congelados, listas suspensas e formatação de status. Uma aba oculta alimenta as listas.
7. **Células vazias mantêm o valor anterior.** Para solicitar exclusão permitida de forma explícita, use `__DELETE__` na célula proposta. A aplicação confere o campo e as dependências antes de aceitar. Conclusão funcional usa `Sim`/`Não`, compatível com diferentes versões do Excel.
8. Use **Import Validated Excel** com os mesmos YAMLs originais carregados. Confira diferenças, registros ausentes, rejeições e conflitos antes de confirmar as propostas aceitas. Identificadores e hashes detectam fontes alteradas; a ordem das linhas não é relevante.
9. Confirme explicitamente a geração do ZIP. **Somente objetos `Validated`, com revisão concluída e sem pendências técnicas, entram no pacote final.** Os relatórios indicam alterações aplicadas, propostas adiadas e objetos ignorados.

Se um arquivo contiver vários objetos e algum não estiver aprovado, o arquivo inteiro é omitido. Não se divide o documento nem se inclui um objeto não aprovado para acomodar os demais. Documentos que contêm apenas catálogos/schemas de referência ficam fora do pacote final de Table/View. O motivo aparece no relatório. Se nada estiver elegível, o ZIP contém somente relatórios.

## Segurança e retenção

- Uploads, YAMLs, ZIPs, planilhas e revisões **não são gravados pela aplicação em disco, banco, armazenamento de objetos ou serviços externos**. Nenhum dado é enviado a APIs de IA ou ao Git.
- O upload Streamlit transmite arquivos do navegador para a memória do servidor Python. Com a configuração padrão, navegador e servidor estão na mesma máquina. Se a configuração for alterada para acesso remoto, os arquivos serão transmitidos ao computador servidor; o processamento não ocorre no navegador.
- O estado fica apenas na memória da sessão Streamlit. Não há cache persistente, `st.cache_data`, `st.cache_resource`, histórico de arquivos ou recuperação de sessões anteriores.
- Reiniciar a sessão limpa os dados e recria os controles de upload. Encerrar a conexão permite ao Streamlit liberar recursos em RAM; não se promete apagamento físico imediato da memória pelo sistema operacional.
- Downloads usam `BytesIO`. Excel usa **XlsxWriter com `in_memory=True`**, inclusive na montagem interna do XLSX; openpyxl lê as planilhas em memória. ZIPs não são extraídos para o sistema de arquivos.
- O usuário decide salvar os downloads. Para retomar uma revisão, deve reabrir os YAMLs originais e o Excel baixado.
- Não há execução de comandos do YAML nem alterações no Unity Catalog. O parser round-trip preserva campos desconhecidos, comentários e ordem sempre que possível.
- Aliases YAML, tags executáveis, chaves duplicadas, profundidade excessiva, links simbólicos, caminhos inseguros, ZIPs aninhados e compactação abusiva são rejeitados.
- Erros do parser são sanitizados. Não se registra conteúdo importado em logs. Detalhes de erro do Streamlit estão desativados para evitar exposição de stack traces.
- O Git ignora YAMLs, planilhas, ZIPs, PDFs, credenciais, ambientes virtuais, configurações pessoais da IDE e logs. Os testes usam somente exemplos sintéticos.

A política descreve o programa. Memória virtual, relatórios de falha, backups e histórico do navegador/sistema operacional não estão sob seu controle.

## Limites configuráveis

Variáveis de ambiente lidas na inicialização; consulte valores padrão e faixas em `app/engine.py`.

| Variável | Finalidade |
|---|---|
| `METADATA_MAX_FILE_BYTES` | Tamanho por YAML |
| `METADATA_MAX_BATCH_BYTES` | Volume descompactado e limite do ZIP de entrada |
| `METADATA_MAX_FILES` | Quantidade de YAMLs no lote |
| `METADATA_MAX_DEPTH` | Profundidade máxima |
| `METADATA_MAX_NODES` | Orçamento de elementos do lote |
| `METADATA_MAX_ZIP_ENTRIES` | Entradas do ZIP, incluindo diretórios |
| `METADATA_MAX_ZIP_RATIO` | Taxa máxima de compactação para entradas grandes |
| `METADATA_MAX_PROCESS_SECONDS` | Prazo cooperativo de importação/geração YAML |

O servidor limita uploads a 10 MiB por arquivo e mensagens a 40 MiB. O serviço de reimportação aceita Excel de até 15 MiB; o uploader pode ser mais restritivo. O prazo é verificado entre operações, sem garantia de interrupção rígida do processo Python. Esta entrega executa localmente; um serviço remoto multiusuário exigiria autenticação e isolamento de recursos.

## Catálogo e referências

Listas e formatos foram transcritos da especificação e conferidos com **DAI-Tags - Databricks Unity Catalog-101026-010433.pdf**. O PDF e o YAML corporativo não são distribuídos.

- O PDF define `[G]`, `[T]` e `[REDUNDANTE]`, mas não atribui categorias a cada linha. Exibe-se **Não identificada**, com pendência global para Governança. Essa lacuna é distinta das inconsistências técnicas dos objetos.
- O PDF apresenta `ct_risco*` em View. O catálogo usa `ct_risco`, conforme a transcrição funcional, e informa a divergência no guia.
- `nm_sist_orig` é aplicável a Table e não é atribuído automaticamente a View.
- Tags são textos, inclusive `"true"` e `"false"`. Ausência nunca se converte em `false` ou `nao_aplicavel`.
- Datas são verificadas no formato `YYYY-MM-DD` e contra o calendário. Grupos, e-mails e `snake_case` têm verificações próprias.
- Classificação considera a coluna mais restritiva e os limites de catálogo/schema. Sem os níveis necessários, fica pendente.
- Não se inventa obrigatoriedade de tags ausentes. Requisitos derivados de mascaramento, classificação e consumo são verificados quando aplicáveis.

## Arquitetura

| Arquivo | Responsabilidade |
|---|---|
| `streamlit_app.py` e `app/ui.py` | Interface e estado da sessão em RAM |
| `app/service.py` | Revisão funcional, comparação e pacote final |
| `app/engine.py` | Parser round-trip, ZIPs e alterações pontuais |
| `app/rules.py` | Catálogo por nível e regras |
| `app/excel.py` | Excel em memória e reimportação |
| `tests/` | Testes com dados sintéticos |

Streamlit atende à tecnologia solicitada. `ruamel.yaml` preserva comentários e estrutura melhor que reconstruir com PyYAML. XlsxWriter evita arquivos temporários no Excel, e openpyxl lê as planilhas preenchidas. Não há build de frontend ou CDN.

**Streamlit está fixado em 1.50.0** porque seu transporte Tornado recebe uploads em memória. O transporte Starlette observado na versão 1.66.1 usa arquivos temporários para uploads maiores que 1 MiB. Não atualize essa dependência sem repetir os testes de transporte. A configuração `disconnectedSessionTTL = 0` desativa a retenção para reconexão; a limpeza explícita revoga também os buffers de upload e download do framework da sessão atual.

## Testes

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -B -m pytest
```

A suíte cobre lotes, ZIPs, round-trip, campos desconhecidos, identidades, classificação, mascaramento, consumo, formatos inválidos, células vazias e exclusões explícitas, conflitos e objetos não aprovados. Testes negam criação de arquivos durante o processamento. O YAML real é verificado apenas em RAM e nunca usado como fixture permanente.

Verificação em 23/10/2026: **111 testes aprovados** com Python 3.12 e Streamlit 1.50.0, incluindo quatro testes da interface e transporte HTTP de upload/download acima de 2 MiB sem escrita em disco. O exemplo corporativo foi processado somente em RAM: uma tabela, 38 colunas, exportação e reimportação sem perda de registros; inconsistências e pendências impediram sua inclusão indevida no pacote aprovado.

## Git

Remoto: `https://github.com/Compacto/metadata-reviewer`; branch padrão confirmada: `main`. Antes de publicar, confira `git status`, execute os testes e revise `git diff --cached --stat`. Versione somente código e documentação. O aplicativo não publica dados no Git ou Databricks.

Documentação: [uploads Streamlit em memória](https://docs.streamlit.io/knowledge-base/using-streamlit/where-file-uploader-store-when-deleted), [configuração Streamlit](https://docs.streamlit.io/develop/api-reference/configuration/config.toml), [opções XlsxWriter](https://xlsxwriter.readthedocs.io/workbook.html).
