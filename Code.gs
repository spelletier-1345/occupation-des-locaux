function makeBilan() {
  
  // Connection au tableur
  var tab = SpreadsheetApp.getActiveSpreadsheet();
    
  // Récupération des données Template
  var tss = SpreadsheetApp.openById("1SV0cyAH18MDf2Xgus6mcNmyjEWU_AkFoXgJ_IoWsbWo")
  var template = tss.getSheetByName("Template");
  template.copyTo(tab);
  
  // Mise en place de la feuille bilan  
  var sheets = tab.getSheets();
  for (var i in sheets) {
    var n = sheets[i].getName();
    if (n == "Bilan") {
      var b = tab.getSheetByName("Bilan");
      tab.deleteSheet(b);}}
  var newTemp = tab.getSheetByName("Copie de Template");
  newTemp.setName("Bilan")
  var bilan = tab.getSheetByName("Bilan");
  var lz = bilan.getRange("B2:B29").getValues()
  zones = []
  for (var x = 0; x < lz.length; x++) {
     zones.push(lz[x][0]);
  }
  var jours = bilan.getRange(1, 6, 1, 10).getValues()[0];
  
  // Feuille résultats
  var feuille = tab.getSheets()[0];
  var lastR = feuille.getLastRow();
  var lastC = feuille.getLastColumn();
  var agents = feuille.getRange(2, 2, lastR, 1).getValues();
  
  // Récupération de la liste de données [agent, jour, zone]
  var datas = [];
  for (var r = 2; r < lastR+2; r++) {
    agent = feuille.getRange(r, 2, 1, 1).getValue();
    for (var c = 3; c < lastC+2; c++) {
      jour = feuille.getRange(1, c, 1, 1).getValue();
      jour = jour.split("]")[0].split("[")[1];
      val = feuille.getRange(r, c, 1, 1).getValue();
      if (val != "") {
        var valList = val.split(", ");
        for (zone in valList) {
          var cell = [agent, jour, valList[zone]];
          datas.push(cell);
        }
      }
    }
  }
  
//  // Vérification des valeurs
//  for (var i = 0; i < datas.length; i++) {
//    var ii = i*4;
//    bilan.getRange(ii+3,17,1,1).setValue(datas[i][0]);
//    bilan.getRange(ii+4,17,1,1).setValue(datas[i][1]);
//    bilan.getRange(ii+5,17,1,1).setValue(datas[i][2]);
//    var lig = zones.indexOf(datas[i][2])
//    bilan.getRange(ii+5,16,1,1).setValue(lig + 2);
//    var col = jours.indexOf(datas[i][1])
//    bilan.getRange(ii+4,16,1,1).setValue(col + 6);
//  }
  
  // Répartition des données dans le tableau
  for (var i = 0; i < datas.length; i++) {
    var lig = zones.indexOf(datas[i][2]) + 2;
    var col = jours.indexOf(datas[i][1]) + 6;
    var cible = bilan.getRange(lig, col, 1, 1)
    var valCell = cible.getValue();
    if (valCell != "") {
      //bilan.getRange(6,1,1,1).setValue(valCell);
      var vc = valCell.split(", ");
      if (vc.indexOf(datas[i][0]) == -1) {
        res = valCell + ", " + datas[i][0]
        cible.setValue(res);
      }
    } else {
      cible.setValue(datas[i][0]);
    }
  }
  
  // Mise en couleur fonction des limites
  var lastR = bilan.getLastRow();
  var lastC = bilan.getLastColumn();
  var jaune = bilan.getRange(1, 1, 1, 1);
  var orange = bilan.getRange(2, 1, 1, 1);
  for (var r = 2; r < lastR+1; r++) {
    var limite = bilan.getRange(r, 4, 1, 1).getValue();
    for (var c = 6; c < lastC+1; c++) {
      var range = bilan.getRange(r, c, 1, 1);
      var occup = range.getValue();
      if (occup != "") {
        occup = occup.split(", ").length;
      } else {
        occup = 0;
      }
      if (occup == limite) {
        jaune.copyFormatToRange(bilan, c, c, r, r);
      } else if (occup > limite) {
        orange.copyFormatToRange(bilan, c, c, r, r);
      }
    }
  }

  bilan.getRange(3, 1, 1, 1).setValue("Done");
}
