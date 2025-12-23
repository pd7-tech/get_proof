# 📘 Manual do Usuário - Extrator de Comprovantes PDF

## 🎯 O que é este programa?

O **Extrator de Comprovantes PDF** é uma ferramenta que automatiza a extração de comprovantes bancários de arquivos PDF grandes, organizando-os por funcionário de acordo com uma planilha Excel.

### Para que serve?

- Você tem um ou vários arquivos PDF com **centenas de comprovantes bancários**
- Você precisa **separar cada comprovante** e salvá-lo individualmente
- Você tem uma **planilha Excel** com os dados dos funcionários (conta, agência, nome)
- O programa **localiza automaticamente** cada comprovante e salva com o nome do funcionário

---

## 📋 Requisitos

### O que você precisa ter:

1. **Windows** (Windows 7 ou superior)
3. **Arquivos necessários:**
   - Pasta com os **PDFs dos comprovantes**
   - **Planilha Excel** (.xlsx ou .xls) com os dados dos funcionários

### Estrutura da Planilha Excel

Sua planilha deve ter **obrigatoriamente** estas colunas (o nome pode variar um pouco):

| Coluna | Exemplos de nomes aceitos | Obrigatória? |
|--------|---------------------------|--------------|
| **Conta** | "Conta", "Conta Corrente", "Account" | ✅ Sim |
| **Agência** | "Agencia", "Agência", "Ag" | ✅ Sim |
| **Nome** | "Nome", "Nome Social", "Funcionario" | ✅ Sim |
| **Centro de Custo** | "Descrição Ccusto", "CCusto", "Setor" | ✅ Sim |

**Exemplo de planilha:**

| Conta | Agencia | Nome Social | Descrição Ccusto |
|-------|---------|-------------|------------------|
| 52938-2 | 0001 | João Silva | TI - Tecnologia |
| 12345-0 | 0001 | Maria Santos | RH - Recursos Humanos |
| 67890-1 | 0002 | Pedro Oliveira | FIN - Financeiro |

> ⚠️ **Importante:** Mantenha os zeros à esquerda na conta e agência (ex: "0001" e não "1")


## ▶️ Como Executar

### Método: Duplo Clique

1. Localize o arquivo `get_proof.py`
2. Dê um **duplo clique** no arquivo
3. A janela do programa irá abrir automaticamente

## 📖 Como Usar o Programa

### Tela Principal

Quando o programa abrir, você verá uma janela com 3 campos principais:

```
┌─────────────────────────────────────────────────┐
│   📁 Arquivos                                   │
├─────────────────────────────────────────────────┤
│ Pasta PDFs:        [___________] [Procurar...]  │
│ Planilha Excel:    [___________] [Procurar...]  │
│ Pasta de Saída:    [___________] [Procurar...]  │
└─────────────────────────────────────────────────┘
```

### Passo 1: Selecionar a Pasta com os PDFs

1. Clique no botão **"Procurar..."** ao lado de "Pasta PDFs"
2. Navegue até a pasta onde estão seus arquivos PDF
3. Selecione a pasta e clique em "Selecionar Pasta"
4. O programa mostrará quantos PDFs foram encontrados

### Passo 2: Selecionar a Planilha Excel

1. Clique no botão **"Procurar..."** ao lado de "Planilha Excel"
2. Navegue até o arquivo Excel (.xlsx ou .xls)
3. Selecione o arquivo e clique em "Abrir"
4. O programa validará automaticamente as colunas

### Passo 3: Escolher a Pasta de Saída

1. Clique no botão **"Procurar..."** ao lado de "Pasta de Saída"
2. Escolha onde os comprovantes extraídos serão salvos
3. Ou deixe o padrão: `comprovantes_extraidos`

> 💡 **Dica:** Se a pasta não existir, o programa irá criá-la automaticamente

### Passo 4: Processar

1. Clique no botão **"▶ PROCESSAR COMPROVANTES"**
2. Aguarde o processamento (um cronômetro mostrará o tempo)
3. Acompanhe o progresso no **Log de Processamento**

---

## 📊 Entendendo os Resultados

### Arquivos Gerados

Os comprovantes são salvos com o seguinte formato de nome:

```
<CentroDeCusto>_<NomeFuncionario>.pdf
```

**Exemplos:**
- `TI-Tecnologia_JoaoSilva.pdf`
- `RH-RecursosHumanos_MariaSantos.pdf`
- `FIN-Financeiro_PedroOliveira.pdf`

### Relatórios Gerados

#### 1. **Log de Processamento** (na tela)

Mostra em tempo real:
- ✓ Comprovantes encontrados e extraídos
- ⚠️ Avisos sobre duplicatas
- ❌ Erros encontrados

#### 2. **Arquivo de Não Encontrados** (se houver)

Se algum comprovante não for localizado, um arquivo TXT será criado:

```
comprovantes_nao_encontrados_<data>_<hora>.txt
```

Este arquivo lista:
- Conta, Nome e Centro de Custo dos não encontrados
- Motivo provável (conta não encontrada, nome diferente, etc.)
- Sugestões para resolver o problema

#### 3. **Histórico de Processamento**

O programa cria um arquivo `pdfs_processados.json` que guarda:
- Quais PDFs já foram processados
- Evita reprocessar os mesmos arquivos

---

## ⚙️ Opções Avançadas

### Ignorar Histórico (Forçar Reprocessamento)

Se você quiser reprocessar PDFs que já foram processados anteriormente:

1. Marque a opção **"Ignorar histórico (forçar reprocessamento)"**
2. Clique em **"▶ PROCESSAR COMPROVANTES"**

### Modo Debug

Para desenvolvedores ou troubleshooting:

1. Marque a opção **"🔧 Debug"**
2. O log mostrará detalhes técnicos da busca

### Limpar Histórico

Para apagar completamente o histórico de PDFs processados:

1. Clique no botão **"🗑️ Limpar Histórico"**
2. Confirme a operação

### Buscar Não Encontrados

Se alguns comprovantes não foram encontrados:

1. Clique em **"🔍 Buscar Não Encontrados"**
2. Escolha entre:
   - **Arquivo TXT**: Carregar lista de não encontrados anterior
   - **Planilha Excel**: Buscar todos os registros novamente
3. Uma nova janela abrirá com busca assistida
4. Selecione um item e clique em **"🔍 Buscar"**
5. O programa tentará localizar com critérios mais flexíveis
6. Se encontrar, clique em **"✓ Extrair Selecionados"**

---

## ❓ Perguntas Frequentes (FAQ)

### 1. O programa não encontrou alguns comprovantes. O que fazer?

**Possíveis causas:**

- **Nome diferente:** O nome no Excel está diferente do nome no PDF
  - **Solução:** Use a busca assistida ou corrija o nome no Excel
  
- **Conta errada:** A conta no Excel não corresponde à conta no PDF
  - **Solução:** Verifique os dígitos da conta e agência
  
- **Comprovante não está no PDF:** O comprovante realmente não existe
  - **Solução:** Verifique se o PDF está completo

### 2. O programa diz que encontrou 0 PDFs, mas eles estão na pasta!

**Solução:**

Se os arquivos estão no **OneDrive ou Google Drive**:

1. Clique com botão direito nos PDFs
2. Escolha **"Sempre manter neste dispositivo"**
3. Aguarde o download completo
4. Ou mova os PDFs para uma pasta local (fora da nuvem)

### 3. O programa está demorando muito. É normal?

**Sim!** O tempo depende de:

- Quantidade de PDFs
- Tamanho dos arquivos
- Quantidade de páginas
- Velocidade do computador

**Estimativa:**
- 10 PDFs com 100 páginas cada = ~2-5 minutos
- 50 PDFs com 500 páginas cada = ~10-20 minutos

### 4. Posso cancelar o processamento?

**Não recomendado**, mas você pode:

1. Fechar a janela do programa (X)
2. Os comprovantes já extraídos serão mantidos
3. Na próxima execução, o programa continuará de onde parou

### 5. O programa extraiu comprovantes duplicados

Isso acontece quando:

- O mesmo comprovante aparece em múltiplas páginas do PDF
- O programa salva todos com sufixo `_dup1`, `_dup2`, etc.

**É normal!** Revise manualmente e delete as duplicatas.

### 6. Erro: "Colunas não encontradas no Excel"

**Solução:**

Verifique se sua planilha tem estas colunas:
- ✅ Conta
- ✅ Agência
- ✅ Nome (ou "Nome Social")
- ✅ Descrição Ccusto (ou "CCusto")

O nome pode ter pequenas variações, mas precisa conter essas palavras.

### 7. Como preservar zeros à esquerda na conta/agência?

**No Excel:**

1. Selecione a coluna de Conta/Agência
2. Clique com botão direito > **"Formatar Células"**
3. Escolha **"Texto"**
4. Digite os valores com zeros à esquerda

Ou adicione um apóstrofo antes: `'0001`

### 8. Posso processar múltiplos lotes de PDFs?

**Sim!** Você tem duas opções:

1. **Colocar todos os PDFs na mesma pasta** e processar de uma vez
2. **Processar pasta por pasta** (o programa mantém histórico)

---

## 🛠️ Solução de Problemas

### Erro: "pip não é reconhecido"

**Causa:** Python não está no PATH

**Solução:**
1. Desinstale o Python
2. Reinstale marcando **"Add Python to PATH"**

### Erro: "ModuleNotFoundError: No module named 'pandas'"

**Causa:** Dependências não instaladas

**Solução:**
```bash
pip install pandas openpyxl xlrd PyPDF2 pdfplumber
```

### Erro ao abrir PDFs grandes

**Causa:** Pouca memória RAM

**Solução:**
1. Feche outros programas
2. Processe em lotes menores (dividir PDFs em pastas)

### Programa trava ou fecha sozinho

**Solução:**
1. Execute pelo Prompt de Comando para ver erros
2. Verifique se os PDFs não estão corrompidos
3. Tente processar um PDF por vez para identificar o problema

---

## 📞 Suporte e Contato

### Precisa de ajuda?

1. Verifique o **arquivo de log** na tela do programa
2. Consulte o **arquivo TXT de não encontrados**
3. Leia novamente este manual

### Reportar Problemas

Ao reportar um problema, inclua:

- ✅ Versão do Windows
- ✅ Mensagem de erro completa
- ✅ Screenshot da tela (se possível)
- ✅ Exemplo de linha do Excel que não funcionou

---

## 📝 Checklist Rápido

Antes de processar, verifique:

- [ ] Planilha Excel com as 4 colunas obrigatórias
- [ ] PDFs na pasta selecionada
- [ ] Pasta de saída escolhida
- [ ] Espaço em disco suficiente

---

## 🎉 Dicas para Melhor Resultado

1. **Organize seus arquivos:**
   - Coloque todos os PDFs em uma única pasta
   - Use nomes de arquivo claros

2. **Mantenha a planilha atualizada:**
   - Corrija dados incorretos antes de processar
   - Use sempre o formato de texto para conta/agência

3. **Primeira execução:**
   - Teste com poucos PDFs primeiro
   - Verifique se os resultados estão corretos
   - Depois processe o lote completo

4. **Backup:**
   - Faça backup dos PDFs originais antes de processar
   - Guarde uma cópia da planilha Excel

5. **Revisão:**
   - Sempre revise os comprovantes extraídos
   - Confira especialmente os casos marcados como duplicados

---

## 📅 Versão

**Versão:** 1.0  
**Última atualização:** Dezembro 2025

---

**💡 Lembre-se:** Este programa foi desenvolvido para facilitar seu trabalho. Se tiver dúvidas, não hesite em ler este manual novamente ou buscar ajuda!

**Bom trabalho! 🚀**
