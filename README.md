# Metadata Review

Aplicativo Streamlit com somente duas funções:

1. **YAML para Excel:** carregue YAMLs/YMLs ou lotes ZIP e baixe uma planilha legível para o Owner.
2. **Excel para YAML:** Governança carrega o Excel devolvido pelo Owner e baixa os YAMLs para revisão e ingestão manual.

Toda revisão acontece no Excel. Não há cadastro, árvore de navegação, aprovação, status obrigatório ou deploy dentro do app.

## Executar no PyCharm

Abra o projeto `metadata-reviewer`, selecione o interpretador `.venv/Scripts/python.exe` e execute `run.py` ou a configuração **Metadata Review**. Acesse **http://127.0.0.1:8501**.

Requer Python 3.11 ou superior; ambiente testado: Python 3.12.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py
```

O servidor escuta apenas no próprio computador. O projeto e as configurações pessoais existentes do PyCharm foram preservados.

## A jornada

Na primeira aba, envie arquivos `.yaml`, `.yml` e/ou `.zip` e clique em **Gerar Excel para o Owner**. A conversão é do lote inteiro; arquivos inválidos e objetos duplicados impedem uma exportação parcial silenciosa.

O Owner abre o Excel, revisa os campos atuais, preenche as alterações desejadas e envia o arquivo à Governança. A planilha tem quatro abas:

| Aba | Uso |
|---|---|
| **Comece aqui** | Guia curto, significado das cores e instruções de preenchimento |
| **Objetos** | Descrições e nomes de negócio |
| **Colunas** | Descrições e informações de privacidade, classificação e mascaramento |
| **Tags** | Tags de objetos e colunas, com explicações em linguagem simples |

**Amarelo** indica proposta de alteração; **azul**, comentário opcional. Os campos atuais são protegidos. Uma proposta em branco mantém o valor original. Para uma remoção explícita, escreva `[REMOVER]`. Listas booleanas usam **Sim/Não**, convertidos para os textos `true`/`false` no YAML. Cabeçalhos e identificadores permanecem congelados, e filtros ajudam a localizar objetos e colunas.

Na aba **Tags**, filtre **Tipo de objeto = Coluna** para revisar as tags colunares. **Objeto** mostra o caminho completo com a tabela e a coluna. São incluídas todas as tags existentes em cada coluna, com as regras do nível Column do PDF. Nenhum valor ausente é inventado. As mesmas tags continuam disponíveis na aba **Colunas**: preencha a proposta em apenas uma aba ou use o mesmo valor nas duas. Propostas diferentes são recusadas com indicação das células em conflito.

Na segunda aba, Governança envia **somente o Excel** e clica em **Gerar YAMLs**. É possível converter mais de um Excel por vez, desde que não tragam versões repetidas dos mesmos objetos. O ZIP contém somente arquivos YAML/YML; se houver um arquivo, também há download individual. O analista confere os arquivos e faz o deploy por seus próprios meios.

Os avisos já existentes na origem não descartam objetos nem exigem aprovação no aplicativo. Valores propostos incompatíveis com formatos/listas documentadas e novas inconsistências são apresentados para correção, sem gerar um pacote parcial. Comentários do Owner aparecem para Governança e não viram campos novos no YAML.

### Mensagem sobre fórmulas

O conversor não calcula fórmulas do Excel. Ele precisa do texto ou valor final a ser aplicado ao YAML. A mensagem identifica a **aba e a célula** afetadas.

No Excel, copie a célula e use **Colar especial → Valores**, salve e envie novamente. Se o conteúdo for um texto literal iniciado por `=`, escreva um apóstrofo antes do sinal. Apenas mudar o formato da célula para Texto não remove uma fórmula existente.

## Excel independente e preservação dos YAMLs

O Excel inclui os arquivos de origem em uma aba técnica oculta. Assim, a segunda etapa funciona sem sessão anterior e **sem reenviar os YAMLs**, preservando comentários, campos desconhecidos, configurações de ingestão, tipos e organização.

Essa aba evita alterações acidentais; **não é criptografia nem assinatura digital**. O Excel deve receber o mesmo cuidado de acesso e compartilhamento dos metadados corporativos originais. Não exclua ou renomeie abas, linhas e colunas do modelo. Identificadores e hashes verificam integridade; não comprovam a identidade do remetente.

Excels da versão anterior sem as origens incorporadas devem ser gerados novamente pela primeira aba. Não se reconstrói um contrato somente a partir das células visíveis, pois isso perderia campos não apresentados ao Owner.

Templates portáteis v1, v2 e v3 são aceitos. A v1 pode mostrar o nome curto da tabela na aba Colunas; a v2 mostra o caminho completo; a v3 também inclui as tags colunares na aba Tags. Planilhas antigas já preenchidas continuam sendo importadas. Para obter as novas linhas de tags de coluna, gere um novo Excel pela primeira aba. Identificadores, hashes e demais campos protegidos continuam sendo conferidos.

## Privacidade

- Uploads, fontes, planilhas, YAMLs gerados e ZIPs são processados **em memória**. O app não os grava em disco, banco de dados ou serviço externo.
- Não há telemetria de uso configurada, APIs de IA, cache persistente ou histórico de revisões.
- O upload é transmitido do navegador à memória do servidor Python. Na configuração entregue, ambos estão no mesmo computador. O processamento não ocorre no navegador.
- O usuário decide salvar e compartilhar os downloads. O arquivo Excel baixado contém os metadados necessários à conversão reversa; isso não é armazenamento pelo servidor.
- **Limpar sessão** descarta estado e buffers de upload/download do Streamlit. A retenção para reconexão está desativada.
- Excel é montado com XlsxWriter `in_memory=True`; leitura e ZIPs usam `BytesIO`. Não há extração para pastas temporárias.
- Streamlit permanece fixado em **1.50.0**, cujo transporte Tornado foi testado com uploads/downloads maiores que 2 MiB sem arquivos temporários. Não atualize essa dependência sem repetir os testes de transporte.
- YAMLs não são executados como código. Caminhos inseguros, aliases, tipos executáveis, ZIPs abusivos, macros, fórmulas e referências externas do Excel são recusados.
- O Git contém somente código, configuração e testes sintéticos. Metadados reais, planilhas, ZIPs, PDFs, credenciais e configurações pessoais da IDE são ignorados.

Memória virtual, backups e relatórios de falha do sistema operacional não estão sob controle do programa.

## Limites e regras

Uploads: até 10 MiB por arquivo; lote de entrada: até 50 MiB. O serviço aceita Excel de até 15 MiB, sujeito ao limite mais restritivo do uploader. Também existem limites cumulativos para fontes recuperadas, arquivos, elementos, volume de saída, profundidade e expansão de ZIPs.

Os limites do parser são configuráveis por variáveis `METADATA_MAX_FILE_BYTES`, `METADATA_MAX_BATCH_BYTES`, `METADATA_MAX_FILES`, `METADATA_MAX_DEPTH`, `METADATA_MAX_NODES`, `METADATA_MAX_ZIP_ENTRIES`, `METADATA_MAX_ZIP_RATIO` e `METADATA_MAX_PROCESS_SECONDS`. Valores e faixas ficam em `app/engine.py`.

O catálogo continua baseado na especificação e no PDF de tags fornecido. Listas externas ausentes e categorias não identificadas são avisadas, sem inventar valores. Diferenças Table/View, strings booleanas, formatos e regras de mascaramento/consumo são respeitados. Não há substituição automática dos arquivos originais nem alteração no Databricks.

## Testes

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -B -m pytest
```

Para incluir os testes com o YAML externo fornecido, sem copiá-lo para o projeto:

```powershell
$env:METADATA_SAMPLE_YAML = 'C:\caminho\para\seu\exemplo.yaml'
.\.venv\Scripts\python.exe -B -m pytest
```

A suíte cobre os dois fluxos, planilha independente da sessão, formatos e integridade, células vazias, exclusões explícitas, fórmulas com indicação de célula, preservação estrutural, ZIPs, ausência de escrita em disco e limpeza de sessão. A verificação real compara o contrato inteiro e testa alterações pontuais e entrada compactada, sempre em RAM.

Verificação em 23/10/2026: **131 testes aprovados**, incluindo compatibilidade dos templates, conflitos entre propostas e cinco testes com o YAML externo fornecido (ida e volta integral, alterações pontuais, entrada ZIP, cobertura das tags colunares e alteração de uma tag na coluna correta). Uma planilha preenchida também foi verificada em memória. Nenhum arquivo corporativo ou resultado desses testes foi salvo no projeto.

## Arquivos principais

- `app/ui.py`: as duas abas Streamlit.
- `app/service.py`: as duas conversões.
- `app/owner_workbook.py`: template legível e fontes portáteis no Excel.
- `app/engine.py`, `app/rules.py`: parser seguro, preservação e regras.
- `app/privacy.py`: limpeza dos buffers da sessão.

Repositório: [Compacto/metadata-reviewer](https://github.com/Compacto/metadata-reviewer), branch `main`. O app não publica arquivos no Git ou no Databricks.
