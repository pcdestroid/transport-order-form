function indiceColuna(x,y,z){
    //exemplo: indiceColuna("texto a procurar","na linha","na Planilha")
    let index = z.getDataRange().getValues()[y - 1].indexOf(x);
    return index+1
  }
  
  function pegarUsuariosTransporte() {
  
    let todosUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuario');
  
    let usuariosTransporte = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users_form_transporte');
  
    var coluna = indiceColuna('Solicitante_transporte',1,todosUsuarios)
    Logger.log(indiceColuna('Solicitante_transporte',1,todosUsuarios))
  
    usuariosTransporte.getRangeList(['A2:A1000']).clear({ contentsOnly: true, skipFilteredRows: true });//Limpar lista
  
    for (let x = 1; x <= todosUsuarios.getLastRow(); x = x + 1) {
      Logger.log(todosUsuarios.getRange(x,1).getValue())
      if (b = todosUsuarios.getRange(x,coluna).getValue() === 'sim')
  
      usuariosTransporte.getRange(usuariosTransporte.getLastRow()+1,1).setValue(todosUsuarios.getRange(x,1).getValue())
      
      }
  
  }
  