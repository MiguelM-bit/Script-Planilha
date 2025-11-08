/**
 * SCRIPT COMPLETO - Combina as duas funcionalidades
 * 
 * 1. Preenche automaticamente a coluna A (DATA) quando novo registro é inserido
 * 2. Converte texto da coluna E (MOTORISTA - CPF - TELEFONE) para MAIÚSCULAS
 */

function onEdit(e) {
  // Obtém informações da edição
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();

  var nomeDaAba = "DADOS"; // Altere para o nome da sua aba
  if (sheet.getName() != nomeDaAba) return;


  var col = range.getColumn();
  
  // Ignora a linha de cabeçalho
  if (row == 1) return;
  
  // FUNCIONALIDADE 1: Preencher data automaticamente na coluna A
  // Quando editar qualquer coluna (exceto A), preenche a data
  if (col != 1) {
    var cellA = sheet.getRange(row, 1);
    
    // Se a célula A estiver vazia, preenche com a data atual
    if (cellA.getValue() == "" || cellA.getValue() == null) {
      // Cria uma data apenas com dia/mês/ano (sem horário)
      var agora = new Date();
      var dataAtual = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate());
      
      cellA.setValue(dataAtual);
      
      // Define o formato de exibição como dd/MM (compacto)
      cellA.setNumberFormat("dd/MM");
    }
  }
  
  // FUNCIONALIDADE 2: Converter coluna E para MAIÚSCULAS
  // Quando editar a coluna E, converte para maiúsculas
  if (col == 5) {
    var valor = range.getValue();
    
    if (valor && typeof valor === 'string') {
      var valorMaiusculo = valor.toUpperCase();
      
      // Só atualiza se for diferente (evita loop infinito)
      if (valor !== valorMaiusculo) {
        range.setValue(valorMaiusculo);
      }
    }
  }
}

/**
 * ===================================================
 * INSTRUÇÕES DE INSTALAÇÃO - VERSÃO SIMPLIFICADA
 * ===================================================
 * 
 * 1. Abra sua planilha Google Sheets
 * 
 * 2. Vá em: Extensions > Apps Script
 * 
 * 3. Delete qualquer código que já estiver lá
 * 
 * 4. Cole TODO este código
 * 
 * 5. Clique em "Salvar" (ícone de disquete ou Ctrl+S)
 * 
 * 6. Configure o Trigger (gatilho):
 *    a) Clique no ícone de RELÓGIO no menu lateral esquerdo
 *    b) Clique em "+ Add Trigger" (botão azul)
 *    c) Configure assim:
 *       - Choose which function to run: onEdit
 *       - Choose which deployment should run: Head
 *       - Select event source: From spreadsheet
 *       - Select event type: On edit
 *    d) Clique em "Save"
 * 
 * 7. Autorize o script:
 *    - Uma janela vai aparecer pedindo permissões
 *    - Clique em "Review permissions"
 *    - Selecione sua conta Google
 *    - Clique em "Advanced" (se aparecer aviso)
 *    - Clique em "Go to [nome do projeto]"
 *    - Clique em "Allow"
 * 
 * 8. Pronto! Agora teste:
 *    - Digite algo em qualquer coluna (B, C, D...)
 *    - A coluna A deve preencher automaticamente com a data
 *    - Digite algo na coluna E
 *    - O texto deve ficar em MAIÚSCULAS automaticamente
 * 
 * ===================================================
 * OBSERVAÇÕES IMPORTANTES
 * ===================================================
 * 
 * - A coluna A (DATA) só será preenchida se estiver vazia
 * - Se você editar a coluna A manualmente, ela não será sobrescrita
 * - A coluna E (MOTORISTA - CPF - TELEFONE) sempre será convertida
 *   para maiúsculas ao digitar
 * - O script funciona para todas as abas da planilha
 * 
 * ===================================================
 * PERSONALIZAÇÕES OPCIONAIS
 * ===================================================
 * 
 * Se quiser que funcione apenas em uma aba específica, adicione:
 * 
 * var nomeDaAba = "Planilha1"; // Altere para o nome da sua aba
 * if (sheet.getName() != nomeDaAba) return;
 * 
 * (adicione essas linhas logo após "var row = range.getRow();")
 */
