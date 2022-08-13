
function enviarEmail() {

    var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base")
    var ul = ultimoLinhaColuna(base, 10) + 1
    if (base.getRange(ul, 1).getValue() != '' & base.getRange(ul, 10).getValue() == '') {
      Logger.log('Enviando email...')
      //################################################################
  
      var solicitanteFrete = base.getRange('B' + ul).getValue()
      var origem = base.getRange('C' + ul).getValue()
      var destino = base.getRange('D' + ul).getValue()
      var pagador = base.getRange('E' + ul).getValue()
      var carga = base.getRange('F' + ul).getValue()
      var descarga = base.getRange('G' + ul).getValue()
      var materiais = base.getRange('H' + ul).getValue()
      materiais = materiais.substring(0, materiais.length - 1).split(",")
      var produtos = []
      var posicao = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
      for (var i = 0; i < (materiais.length) / 10; i = i + 1) {
  
        let prod = materiais[posicao[0]];
        let qtd = materiais[posicao[1]];
        let um = materiais[posicao[2]];
        let valor = materiais[posicao[3]];
        let peso = materiais[posicao[4]];
        let comp = materiais[posicao[5]];
        let larg = materiais[posicao[6]];
        let alt = materiais[posicao[7]];
        let vol = materiais[posicao[8]];
        let tria = materiais[posicao[9]];
  
        produtos[i] = {}; produtos[i].produto = prod; produtos[i].quantidade = qtd;
        produtos[i].valor = valor; produtos[i].um = um; produtos[i].peso = peso;
        produtos[i].comp = comp; produtos[i].larg = larg; produtos[i].alt = alt;
        produtos[i].vol = vol; produtos[i].tria = tria;
  
        for (let x = 0; x < 10; x = x + 1) { posicao[x] = posicao[x] + 10; }
  
      }
  
      var x = carga.split("-"); carga = x[2] + '/' + x[1] + '/' + x[0]
      var y = descarga.split("-"); descarga = y[2] + '/' + y[1] + '/' + y[0]
      var listaProdutoshtml = '';
      for (var i = 0; i < (produtos.length); i = i + 1) {
        
        listaProdutoshtml = listaProdutoshtml + '<tr style="height: 18.4px;">';
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].produto +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].quantidade.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].valor.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].peso.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].comp.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].larg.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].alt.replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ produtos[i].vol.replace(".", ",") +'</td>'
        x = (produtos[i].comp * produtos[i].larg * produtos[i].alt)
        y = produtos[i].vol
        listaProdutoshtml = listaProdutoshtml + '<td style="height: 18.4px;">'+ (x*y).toString().replace(".", ",") +'</td>'
        listaProdutoshtml = listaProdutoshtml + '</tr>';
  
      }
      //################################################################
  
      var email = ''
      var subject = "(Teste) - Pedido de transporte - " + origem + " x " + destino
      var assinatura = ''
  
      var emailTemp = HtmlService.createHtmlOutput('<span>A pedido de ' + solicitanteFrete + ' foi solicitado um serviço de transporte com origem na ' + origem + ' e destino na ' + destino + '.</span><br><span>Carregamento previsto no dia ' + carga + ' e descarregamento previsto para o dia ' + descarga + '</span><br><span>O pagamento do frete ficará por conta da ' + pagador + '.</span><br><br><span>Segue produtos abaixo:</span><br><br>' + '<table style="border-collapse: collapse; text-align: center; width: 100%; height: 36.8px;" border="1"><colgroup><col style="width: 26%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9%;"><col style="width: 9.48181%;"></colgroup><tbody><tr style="height: 18.4px;"><td style="height: 18.4px;"><strong>Produto</strong></td><td style="height: 18.4px;"><strong>Qtd.</strong></td><td style="height: 18.4px;"><strong>Valor</strong></td><td style="height: 18.4px;"><strong>Peso(kg)</strong></td><td style="height: 18.4px;"><strong>Comp(m)</strong></td><td style="height: 18.4px;"><strong>Largura(m)</strong></td><td style="height: 18.4px;"><strong>Altura(m)</strong></td><td style="height: 18.4px;"><strong>Volume</strong></td><td style="height: 18.4px;"><strong>Volumetria</strong></td></tr>' + listaProdutoshtml +'</tbody></table><br><br><span>Atenciosamente,</span><br><span><strong>' + assinatura + '</strong></span><br><p><img src="https://intranet.brmarinas.com.br/wp-content/uploads/2020/05/logo.png" width="136" height="48"></p>').setTitle('Email');
  
      var htmlMessage = emailTemp.getContent()
      //Logger.log(htmlMessage);
  
      GmailApp.sendEmail(email, subject, "You email does't support HTML.", { name: 'Pedido de transporte', htmlBody: htmlMessage });
      base.getRange(ul, 10).setValue('Ok');
      Logger.log('Email enviado!')
    }
  
  }
  
  //#######################################################################
  
  function concluirPedido() {
    var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base")
    if (base.getRange(base.getLastRow(), 9).getValue() == "") {
      //Pegar última  linha preenchida da coluna
      var coluna = 9; // Qual coluna?
      x = 1; do { console.log(base.getRange(x, coluna).getValue()); x++; }
      while (base.getRange(x, coluna).getValue() != "");
      base.getRange(x, coluna).setValue('Ok')
      novoPedido()
    }
  
  }
  
  //#######################################################################
  
  function novoPedido() {
    var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base")
  
    if (base.getRange(base.getLastRow(), 9).getValue() == "") {
  
      var pedidoRecente = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pedido recente")
  
      //Pegar última  linha preenchida da coluna
      var coluna = 9; // Qual coluna?
      x = 1; do { console.log(base.getRange(x, coluna).getValue()); x++; }
      while (base.getRange(x, coluna).getValue() != "");
      var ul = x
      console.log(ul)
  
      let d = base.getRange('A' + ul).getValue()
      var dataRegistro = zeroEsquerda(d.getDate(), 2) + '-' + zeroEsquerda(d.getMonth(), 2) + '-' + d.getFullYear()
  
      Logger.log(dataRegistro)
  
      var solicitanteFrete = base.getRange('B' + ul).getValue()
      var origem = base.getRange('C' + ul).getValue()
      var destino = base.getRange('D' + ul).getValue()
      var pagador = base.getRange('E' + ul).getValue()
      var carga = base.getRange('F' + ul).getValue()
      var descarga = base.getRange('G' + ul).getValue()
      var materiais = base.getRange('H' + ul).getValue()
  
      materiais = materiais.substring(0, materiais.length - 1).split(",")
  
      var produtos = []
  
      var posicao = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
      for (var i = 0; i < (materiais.length) / 10; i = i + 1) {
  
        let prod = materiais[posicao[0]];
        let qtd = materiais[posicao[1]];
        let um = materiais[posicao[2]];
        let valor = materiais[posicao[3]];
        let peso = materiais[posicao[4]];
        let comp = materiais[posicao[5]];
        let larg = materiais[posicao[6]];
        let alt = materiais[posicao[7]];
        let vol = materiais[posicao[8]];
        let tria = materiais[posicao[9]];
  
        produtos[i] = {}; produtos[i].produto = prod; produtos[i].quantidade = qtd;
        produtos[i].valor = valor; produtos[i].um = um; produtos[i].peso = peso;
        produtos[i].comp = comp; produtos[i].larg = larg; produtos[i].alt = alt;
        produtos[i].vol = vol; produtos[i].tria = tria;
  
        for (let x = 0; x < 10; x = x + 1) { posicao[x] = posicao[x] + 10; }
  
      }
  
      //Inserir produtos na lista
      ul = 10;
      for (var i = 0; i < (produtos.length); i = i + 1) {
        pedidoRecente.getRange('A' + (ul + i + 1)).setValue(i + 1)
        Logger.log(i)
        pedidoRecente.getRange('C' + (ul + i + 1)).setValue(produtos[i].produto)
        pedidoRecente.getRange('B' + (ul + i + 1)).setValue(produtos[i].quantidade.replace(".", ","))
        pedidoRecente.getRange('D' + (ul + i + 1)).setValue(produtos[i].valor.replace(".", ","))
        pedidoRecente.getRange('E' + (ul + i + 1)).setValue(produtos[i].peso.replace(".", ","))
        pedidoRecente.getRange('F' + (ul + i + 1)).setValue(produtos[i].comp.replace(".", ","))
        pedidoRecente.getRange('G' + (ul + i + 1)).setValue(produtos[i].larg.replace(".", ","))
        pedidoRecente.getRange('H' + (ul + i + 1)).setValue(produtos[i].alt.replace(".", ","))
        pedidoRecente.getRange('I' + (ul + i + 1)).setValue(produtos[i].vol.replace(".", ","))
        pedidoRecente.getRange('J' + (ul + i + 1)).setValue(((produtos[i].comp * produtos[i].larg * produtos[i].alt) * produtos[i].vol).toString().replace(".", ","))
  
        pedidoRecente.getRangeList(['A' + (ul + i + 2) + ':J1000']).clear({ contentsOnly: true, skipFilteredRows: true });//Limpar produtos da lista
      }
  
      var x = carga.split("-")
      carga = x[2] + '-' + x[1] + '-' + x[0]
      var y = descarga.split("-")
      descarga = y[2] + '-' + y[1] + '-' + y[0]
  
      var dados = [solicitanteFrete, origem, destino, pagador, carga, descarga]
  
      for (var i = 0; i < (dados.length); i = i + 1) {
        pedidoRecente.getRange('B' + (3 + i)).setValue(dados[i].toString().toUpperCase())
      }
      pedidoRecente.getRange('B2').setValue(dataRegistro)
    }
  }
  function zeroEsquerda(value, totalWidth, paddingChar) {
    var length = totalWidth - value.toString().length + 1;
    return Array(length).join(paddingChar || '0') + value;
  };
  function ultimoValorColuna(planilha, coluna) {
    //Pegar o último valor de uma coluna específica.
    x = 1; do { ; x++; }
    while (planilha.getRange(x, coluna).getValue() != "");
    return planilha.getRange(x - 1, coluna).getValue()
  }
  
  function ultimoLinhaColuna(planilha, coluna) {
    //Pegar última linha de uma coluna específica.
    x = 1; do { ; x++; }
    while (planilha.getRange(x, coluna).getValue() != "");
    return (x - 1)
  }



