# 📊 Google Sheets Auto-Scripts

![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
![Apps Script](https://img.shields.io/badge/Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-blue.svg?style=for-the-badge)

Scripts de automação inteligente para Google Sheets que otimizam o preenchimento de dados e garantem padronização de informações.

## 🎯 O Que Este Script Faz?

Automatiza duas tarefas essenciais em planilhas Google Sheets:

### ✅ **1. Data Automática (Coluna A)**
Registra automaticamente a data de criação de cada novo registro, sem necessidade de preenchimento manual.

**Como funciona:**
- 📅 Quando você adiciona dados em **qualquer coluna** (B, C, D, E, F...)
- 🤖 A **Coluna A** é preenchida automaticamente com a data atual
- 🎨 **Visualização compacta**: `07/11` (dia/mês)
- 💾 **Valor armazenado**: `07/11/2025` (data completa para cálculos)
- ⏱️ **Sem horário**: Apenas a data, sem horas/minutos/segundos

### ✅ **2. Padronização de Texto em Maiúsculas (Coluna E)**
Converte automaticamente textos para MAIÚSCULAS, garantindo uniformidade nos dados.

**Como funciona:**
- ⌨️ Quando você digita na **Coluna E** (MOTORISTA - CPF - TELEFONE)
- 🔄 O texto é convertido instantaneamente para **MAIÚSCULAS**
- 📝 Exemplo: `joão silva - 123.456.789-00` → `JOÃO SILVA - 123.456.789-00`
- ⚡ Conversão em tempo real enquanto você digita

## 🚀 Instalação Rápida (5 minutos)

### **Passo 1: Abrir o Editor de Scripts**
1. Abra sua planilha no Google Sheets
2. Clique em **Extensions** → **Apps Script**

### **Passo 2: Adicionar o Código**
1. Delete qualquer código existente no editor
2. Copie **todo** o conteúdo do arquivo [`scriptCompleto.gs`](scriptCompleto.gs)
3. Cole no editor do Apps Script
4. Clique em **💾 Salvar** (`Ctrl+S`)

### **Passo 3: Configurar o Trigger (Gatilho)**
1. No Apps Script, clique no ícone de **⏰ Relógio** (Triggers) no menu lateral
2. Clique em **➕ Add Trigger**
3. Configure assim:
   ```
   Function to run: onEdit
   Deployment: Head
   Event source: From spreadsheet
   Event type: On edit
   ```
4. Clique em **Save**

### **Passo 4: Autorizar Permissões**
1. Uma janela solicitará permissões
2. Clique em **Review permissions**
3. Selecione sua conta Google
4. Se aparecer aviso de segurança:
   - Clique em **Advanced**
   - Clique em **Go to [nome do projeto] (unsafe)**
5. Clique em **Allow**

### **Passo 5: Testar! 🎉**
- Digite algo em qualquer coluna (B, C, D...) → Coluna A preenche automaticamente ✅
- Digite texto na coluna E → Vira MAIÚSCULAS automaticamente ✅

## 📁 Estrutura do Repositório

```
📦 macro/
├── 📄 scriptCompleto.gs              # ⭐ Script principal (USE ESTE!)
├── 📄 preencherDataAutomatica.gs     # Script individual - Data automática
├── 📄 colunaEMaiusculo.gs            # Script individual - Maiúsculas
└── 📄 README.md                      # Documentação completa
```

**Recomendação:** Use o **`scriptCompleto.gs`** para ter ambas as funcionalidades ativas.

## ⚙️ Personalização Avançada

### 🎯 Aplicar apenas em uma aba específica

Por padrão, o script funciona na aba **"DADOS"**. Para alterar:

```javascript
var nomeDaAba = "DADOS"; // Altere para o nome da sua aba
if (sheet.getName() != nomeDaAba) return;
```

Localize estas linhas no início da função `onEdit` e altere `"DADOS"` para o nome da sua aba.

### 📅 Alterar formato de exibição da data

Localize a linha:
```javascript
cellA.setNumberFormat("dd/MM");
```

Substitua por:
| Formato | Exibição | Exemplo |
|---------|----------|---------|
| `"dd/MM"` | Dia/Mês (compacto) | `07/11` |
| `"dd/MM/yyyy"` | Completo | `07/11/2025` |
| `"dd/MM/yy"` | Ano curto | `07/11/25` |
| `"yyyy-MM-dd"` | Internacional | `2025-11-07` |
| `"dd/MM/yyyy HH:mm"` | Com hora | `07/11/2025 14:30` |

### 📝 Aplicar maiúsculas em outra coluna

Para converter outra coluna além da E, altere:
```javascript
if (col == 5) { // 5 = Coluna E
```

Tabela de referência:
| Coluna | Número |
|--------|--------|
| A | 1 |
| B | 2 |
| C | 3 |
| D | 4 |
| E | 5 |
| F | 6 |
| ... | ... |

### 🔄 Permitir múltiplas colunas em maiúsculas

```javascript
if (col == 5 || col == 7 || col == 9) { // Colunas E, G, I
  // ... código de conversão
}
```

## 💡 Comportamento e Características

### ✅ **Data Automática (Coluna A)**
- ✔️ **Só preenche se estiver vazia** - Não sobrescreve datas existentes
- ✔️ **Edição manual preservada** - Se você editar a data manualmente, ela não será alterada
- ✔️ **Sem horário** - Registra apenas dia/mês/ano (sem horas/minutos/segundos)
- ✔️ **Formato compacto** - Exibe `dd/MM` mas armazena data completa
- ✔️ **Ignora linha de cabeçalho** - Não preenche a linha 1

### 🔤 **Conversão para Maiúsculas (Coluna E)**
- ✔️ **Conversão instantânea** - Acontece em tempo real enquanto você digita
- ✔️ **Sempre ativa** - Toda edição na coluna E é convertida
- ✔️ **Preserva números e símbolos** - Apenas letras são convertidas
- ✔️ **Evita loops infinitos** - Só converte quando necessário

### 🎯 **Configuração Atual**
- 📍 **Aba ativa:** `DADOS` (pode ser alterada - veja seção Personalização)
- 📍 **Coluna de data:** A (primeira coluna)
- 📍 **Coluna de maiúsculas:** E (quinta coluna)

## 🐛 Solução de Problemas

<details>
<summary><b>❌ O script não está funcionando</b></summary>

**Possíveis causas e soluções:**

1. **Trigger não configurado**
   - Verifique se o trigger foi criado corretamente
   - Vá em Apps Script → ⏰ Triggers e confira se `onEdit` está listado

2. **Permissões não concedidas**
   - Execute o script manualmente uma vez
   - Autorize quando solicitado

3. **Nome da aba incorreto**
   - O script está configurado para a aba `"DADOS"`
   - Verifique se sua aba tem exatamente este nome
   - Ou altere o nome no código (veja seção Personalização)

4. **Verificar erros no log**
   - Apps Script → View → **Execution log**
   - Verifique mensagens de erro

</details>

<details>
<summary><b>📅 A data está aparecendo errada</b></summary>

**Se a data estiver com horário:**
- Certifique-se de usar a versão mais recente do `scriptCompleto.gs`
- O código deve ter a linha: `new Date(agora.getFullYear(), agora.getMonth(), agora.getDate())`

**Se o formato estiver errado:**
- Selecione a coluna A
- Vá em `Format` → `Number` → `Custom date and time`
- Digite: `dd/MM` ou `dd/MM/yyyy`

</details>

<details>
<summary><b>🔤 A conversão para maiúsculas não funciona</b></summary>

**Verificações:**
1. Certifique-se de estar editando a **Coluna E**
2. Não funciona com números (apenas texto)
3. Aguarde um instante após digitar (conversão é automática)

</details>

<details>
<summary><b>⚠️ Script funciona em todas as abas</b></summary>

Se quiser que funcione apenas em uma aba específica:
1. Localize as linhas no código:
   ```javascript
   var nomeDaAba = "DADOS";
   if (sheet.getName() != nomeDaAba) return;
   ```
2. Altere `"DADOS"` para o nome da sua aba

</details>

<details>
<summary><b>🔧 Como ver o log de execução</b></summary>

1. Abra o Apps Script editor
2. Clique em `View` → `Execution log`
3. Ou clique em `View` → `Logs` (Ctrl + Enter)
4. Edite algo na planilha para gerar logs

</details>

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:
- 🐛 Reportar bugs
- 💡 Sugerir novas funcionalidades
- 🔧 Enviar pull requests
- ⭐ Dar uma estrela no repositório

## � Licença

Este projeto está sob a licença MIT. Livre para uso, modificação e distribuição.

---

## 📬 Contato & Suporte

- 🌟 **Deixe uma estrela** se este script foi útil!
- 🐛 **Abra uma issue** para reportar problemas
- 💬 **Discussões** para tirar dúvidas ou compartilhar ideias

---

<div align="center">

**Desenvolvido por Miguel Matheus para automatizar sua vida no Google Sheets**

[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=flat&logo=google-sheets&logoColor=white)](https://sheets.google.com)
[![Apps Script](https://img.shields.io/badge/Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)](https://developers.google.com/apps-script)

</div>
