//***Globals***
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sr = ss.getSheetByName("RANKING"); //Pestaña RANKING
var playersRankingRows = { //Matriz de resultados jugadores / parejas en la pestaña RANKING
  "AN" : 89, "GA": 90, "HE": 91, "HU" : 92, "JO" : 93, "LA" : 94,
  "AN-GA" : 104, "GA-AN":104, "AN-HE" : 105, "HE-AN" : 105, "AN-HU" : 106, "HU-AN" : 106,
  "AN-JO" : 107, "JO-AN":107, "AN-LA" : 108, "LA-AN" : 108, "GA-HE" : 109, "HE-GA" : 109,
  "GA-HU" : 110, "HU-GA":110, "GA-JO" : 111, "JO-GA" : 111, "GA-LA" : 112, "LA-GA" : 112,
  "HE-HU" : 113, "HU-HE":113, "HE-JO" : 114, "JO-HE" : 114, "HE-LA" : 115, "LA-HE" : 115,
  "HU-JO" : 116, "JO-HU":116, "HU-LA" : 117, "LA-HU" : 117, "JO-LA" : 118, "LA-JO" : 118
}

var playersRankingGraphics = { //Matriz rankings
  "AN" : 1, "GA": 2, "HE": 3, "HU" : 4, "JO" : 5, "LA" : 6,
  "AN-GA" : 8, "GA-AN":8, "AN-HE" : 9, "HE-AN" : 9, "AN-HU" : 10, "HU-AN" : 10,
  "AN-JO" : 11, "JO-AN":11, "AN-LA" : 12, "LA-AN" : 12, "GA-HE" : 13, "HE-GA" : 13,
  "GA-HU" : 14, "HU-GA":14, "GA-JO" : 15, "JO-GA" : 15, "GA-LA" : 16, "LA-GA" : 16,
  "HE-HU" : 17, "HU-HE":17, "HE-JO" : 18, "JO-HE" : 18, "HE-LA" : 19, "LA-HE" : 19,
  "HU-JO" : 20, "JO-HU":20, "HU-LA" : 21, "LA-HU" : 21, "JO-LA" : 22, "LA-JO" : 22
}

var rangeGraphics = sr.getRange("C122:Y622").getValues(); //Ranking máximo de 500 partidos

var rankingColumn = "D"; //Columna Ranking en la pestaña RANKING
var matchWonRankingColumn = "F" //Columna de partidos ganados en la pestaña RANKING
var matchLostRankingColumn = "G"; //Columna de partidos perdidos en la pestaña RANKING
var gameWonRankingColumn = "H"; //Columna de juegos ganados en la pestaña RANKING
var gameLostRankingColumn = "I"; //Columna de juegos perdidos en la pestaña RANKING


var rowStart = 3; //Comienzo de las celdas de partidos
var rowIncrement = 4; //Salto a los siguientes tres partidos
var columnsResult = ["B","F","H","L","N","R"] //Columnas inicio fin de los partidos
var columnsResultGames = ["D","E","J","K","P","Q"] //Columnas para obtener los resultados de los partidos

var resultStyleOK = "#b5d7a8"; //color verde
var resultStyleKO = "#f4cccc"; //color rojo


var alignmentsRange = "Z2"; //Constantes para cargar los jugadores



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Padel')
      .addItem('Cargar Jugadores', "loadPlayers")
      .addSeparator()
      .addItem('Guardar Resultados', 'saveResults')
      .addToUi();
}

/**
 * Función para cargar los jugadores en las tablas "partido" del excel, desde una celda con el siguiente formato en texto plano:
 * [local1-local2#visitante1-visitante2, local1-local2#visitante1-visitante2 ....]
 */
function loadPlayers(){
  var players = ss.getRange(alignmentsRange).getValue(); //Obtiene las alineaciones desde Z2
  if(players === ""){ //comprueba que existan jugadores
    showAlert("No hay jugadores para cargar");
    return;
  }

  var matches = players.substring(1, players.length-1).split(","); //realiza un split de cada uno de los partidos
  var limit = rowStart + 15*rowIncrement;  //15 filas, 45 partidos, 6 jugadores
  var matchIndex = 0; //indice para seleccionar los partidos

  for(var i=rowStart; i < limit;i=i+rowIncrement){ //filas de partidos
    for(var j=0;j<6;j=j+2){ //columna
      var matchRange = ss.getRange(columnsResult[j]+i+":"+columnsResult[j+1]+(i+1));
      var matchRangeValues = matchRange.getValues(); //Rango de celdas de un partido
      if(matchRangeValues[0][0] === ""){ //cuando ya no haya mas "cajas" de partido termina el bucle
        i = limit;
        break;
      }
      var matchWithPlayersValues = loadMatch(matchRangeValues, matches[matchIndex]); //carga los valores con los nombres de los jugadores del partido
      matchRange.setValues(matchWithPlayersValues); //Pone los valores en el excel, en su correspondiente partido

      matchIndex = matchIndex + 1;
    }
  }

  ss.getRange(alignmentsRange).setValue(""); //elimina los valores de los jugadores en "plano"
  showAlert("Jugadores Cargados!");
}

function loadMatch(matchRange,match){
  match = match.trim(); //elimina espacios por delante y por detrás del partido
  var pairs = match.split("#"); //separa la información en pareja local/pareja visitante
  var playersLocal = pairs[0].split("-"); //separa la pareja local en jugador local 1 y jugador local 2
  var playersVisitor = pairs[1].split("-"); //separa la pareja visitante en jugador visitante 1 y jugador visitante 2
  matchRange[0][1] = playersLocal[0]; //Pone los valores en la matriz
  matchRange[1][1] = playersLocal[1];
  matchRange[0][4] = playersVisitor[0];
  matchRange[1][4] = playersVisitor[1];
  return matchRange;
}


/**
 * Itera sobre cada tabla partido del excel.
 * Filas de máximo hasta 45 partidos (15 filas)
 * Columnas de tres partidos por fila
 */
function saveResults() {
  if(ss.getRange("A1").getValue()==="*"){
    showAlert("Los resultados ya están añadidos!");
    return;
  }
  var limit = rowStart + 15*rowIncrement;  //15 filas, 45 partidos, 6 jugadores

  for(var i=rowStart; i < limit;i=i+rowIncrement){ //filas de partidos
    for(var j=0;j<6;j=j+2){ //columna
      var matchRange = ss.getRange(columnsResult[j]+i+":"+columnsResult[j+1]+(i+1)).getValues(); //Rango de celdas de un partido
      if(matchRange[0][0] === ""){
        i = limit;
        break;
      }
      writeMatch(matchRange,i, columnsResultGames[j], columnsResultGames[j+1]); //escribe el partido
    }
  }
  sr.getRange("C122:Y622").setValues(rangeGraphics); //actualizamos los resultados
  ss.getRange("A1").setValue("*");
  showAlert("Resultados Añadidos!");

}
/**
 * matchRange: son los valores de la tabla que contiene un partido (nombres de jugadores y resultados)
 * row: es la fila actual en la que estamos
 * columnLocal: es la columna de resultados del equipo local dentro de ese partido
 * columnVisit: es la columna de resultados del equipo visitante dentro de ese partido
 */
function writeMatch(matchRange, row, columnLocal, columnVisit){
  var localPairArray = convertToPlayersPair(matchRange[0][1], matchRange[1][1]); //Convierte los nombres del equipo local a prefijos
  var visitPairArray = convertToPlayersPair(matchRange[0][4], matchRange[1][4]); //Convierte los nombres del equipo visitante a prefijos
  var localResult = matchRange[0][2]; //Obtiene el resultado en juegos del equipo local
  var visitResult = matchRange[0][3]; //Obitene el resultado en juegos del equipo visitante
  var rankingResultIndividuals;
  var rankingResultPair;

  if(localResult !== "" && visitResult !== ""){ //Si no hay resultados salta el partido
    addResult(localPairArray[0], gameWonRankingColumn, parseInt(localResult)); //añadir juegos ganados al jugador local 1
    addResult(localPairArray[1], gameWonRankingColumn, parseInt(localResult)); //añadir juegos ganados al jugador local 2
    addResult(localPairArray[2], gameWonRankingColumn, parseInt(localResult)); //añadir juegos ganados a la pareja local
    addResult(localPairArray[0], gameLostRankingColumn, parseInt(visitResult)); //añadir juegos perdidos al jugador local 1
    addResult(localPairArray[1], gameLostRankingColumn, parseInt(visitResult)); //añadir juegos perdidos al jugador local 2
    addResult(localPairArray[2], gameLostRankingColumn, parseInt(visitResult)); //añadir juegos perdidos a la pareja local

    addResult(visitPairArray[0], gameWonRankingColumn, parseInt(visitResult)); //añadir juegos ganados al jugador visitante 1
    addResult(visitPairArray[1], gameWonRankingColumn, parseInt(visitResult)); //añadir juegos ganados al jugador visitante 2
    addResult(visitPairArray[2], gameWonRankingColumn, parseInt(visitResult)); //añadir juegos ganados a la pareja visitante
    addResult(visitPairArray[0], gameLostRankingColumn, parseInt(localResult)); //añadir juegos perdidos al jugador visitante 1
    addResult(visitPairArray[1], gameLostRankingColumn, parseInt(localResult)); //añadir juegos perdidos al jugador visitante 2
    addResult(visitPairArray[2], gameLostRankingColumn, parseInt(localResult)); //añadir juegos perdidos a la pareja visitante

    if(parseInt(localResult) > parseInt(visitResult)){ // si ha ganado la pareja local
      addResult(localPairArray[0], matchWonRankingColumn, 1); //añadir partido ganado al jugador local 1
      addResult(localPairArray[1], matchWonRankingColumn, 1); //añadir partido ganado al jugador local 2
      addResult(localPairArray[2], matchWonRankingColumn, 1); //añadir partido ganado a la pareja local
      addResult(visitPairArray[0], matchLostRankingColumn, 1); //añadir partido perdido al jugador visitante 1
      addResult(visitPairArray[1], matchLostRankingColumn, 1); //añadir partido perdido al jugador visitante 2
      addResult(visitPairArray[2], matchLostRankingColumn, 1); //añadir partido perdido a la pareja visitante
      rankingResultPair = writePairRanking(localPairArray[2], visitPairArray[2]); //añadir ranking a las parejas siendo ganador el equipo local
      rankingResultIndividuals = writeInvidualRanking(localPairArray[0],localPairArray[1], visitPairArray[0], visitPairArray[1]); //añadir ranking a los individuos
      writeChartRankings(localPairArray,visitPairArray,rankingResultIndividuals,rankingResultPair, true);
      ss.getRange(columnLocal+row).setBackground(resultStyleOK); //Pone un color verde al resultado del equipo local
      ss.getRange(columnVisit+row).setBackground(resultStyleKO); //Pone un color rojo al resultado del equipo visitante
    }else{ //si ha ganado la pareja visitante
      addResult(localPairArray[0], matchLostRankingColumn, 1); //añade partido perdido al jugador local 1
      addResult(localPairArray[1], matchLostRankingColumn, 1); //añade partido perdido al jugador local 2
      addResult(localPairArray[2], matchLostRankingColumn, 1); //añade partido perdido a la pareja local
      addResult(visitPairArray[0], matchWonRankingColumn, 1); //añade partido ganado al jugador visitante 1
      addResult(visitPairArray[1], matchWonRankingColumn, 1); //añade partido ganado al jugador visitante 2
      addResult(visitPairArray[2], matchWonRankingColumn, 1); //añade partido ganado  la pareja visitante
      rankingResultPair = writePairRanking(visitPairArray[2], localPairArray[2]); //añadir ranking a las parejas siendo ganador el equipo visitante
      rankingResultIndividuals = writeInvidualRanking(visitPairArray[0],visitPairArray[1], localPairArray[0], localPairArray[1]); //añadir ranking a los individuos
      writeChartRankings(localPairArray,visitPairArray,rankingResultIndividuals,rankingResultPair, false);
      ss.getRange(columnLocal+row).setBackground(resultStyleKO); //Pone un color rojo al resultado del equipo local
      ss.getRange(columnVisit+row).setBackground(resultStyleOK); //Pone un color verde al resultado del equipo visitante
    }

  }
}
/**
 * player: es el prefijo o el de un jugador o una pareja, por ejemplo: Huertas -> HU, pareja Huertas-Garcho -> HU-GA
 * resultColumn: es la columna de resultados correspondiente a la pestaña 'RANKING' (puede ser la columna de juegos ganados, perdidos o partido ganado o perdido)
 * resultValue: son los juegos ganados o los partidos ganados o perdidos
 */

function addResult(player, resultColumn, resultValue){
  var actualResult = parseInt(sr.getRange(resultColumn+playersRankingRows[player]).getValue());
  var updatedResult = actualResult + resultValue;
  sr.getRange(resultColumn+playersRankingRows[player]).setValue(updatedResult);
}

/** convertToPlayersPair
 * Convierte nombres completos a prefijos, por ejemplo
 * Huertas-Garcho a HU-GA
 * devuelve -
 * localPairArray[0] -> jugador1 local (Ejemplo: GA -> Garcho)
 * localPairArray[1] -> jugador2 local (Ejemplo: HU -> Huertas)
 * localPairArray[2] -> pareja local (Ejemplo: GA-HU)
 */
function convertToPlayersPair(player1, player2){
  var prefixPlayer1 = player1.toUpperCase().replace("É", "E").substring(0,2);
  var prefixPlayer2 = player2.toUpperCase().replace("É", "E").substring(0,2);
  return [prefixPlayer1, prefixPlayer2, prefixPlayer1+"-"+prefixPlayer2];
}

/**
 * playerWinner1: prefijo del jugador ganador 1 (Ejemplo Huertas -> HU)
 * playerWinner2: prefijo del jugador ganador 2 (Ejemplo Huertas -> HU)
 * playerLooser1: prefijo del jugador perdedor 1 (Ejemplo Huertas -> HU)
 * playerLooser2: prefijo del jugador perdedor 2 (Ejemplo Huertas -> HU)
 */

function writeInvidualRanking(playerWinner1, playerWinner2, playerLooser1, playerLooser2){
  var rankingWinner1 = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerWinner1]).getValue()); //obtiene el ranking actual del jugador ganador 1
  var rankingWinner2 = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerWinner2]).getValue()); //obtiene el ranking actual del jugador ganador 2
  var rankingLooser1 = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerLooser1]).getValue()); //obtiene el ranking actual del jugador perdedor 1
  var rankingLooser2 = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerLooser2]).getValue()); //obtiene el ranking actual del jugador perdedor 2
  var rankingWinner = (rankingWinner1+rankingWinner2)/2; //Calcula una media de los rankings de los jugadores ganadores
  var rankingLooser = (rankingLooser1+rankingLooser2)/2; //Calcula una media de los rankings de los jugadores perdedores
  var rankingVariation = formuleRanking(rankingWinner, rankingLooser); //Calcula la variación del ranking

  rankingWinner1 = rankingWinner1 + rankingVariation; //suma de la variación de ranking a los jugadores ganadores
  rankingWinner2 = rankingWinner2 + rankingVariation;
  rankingLooser1 = rankingLooser1 - rankingVariation; //resta de la variación de ranking a los jugadores pededores
  rankingLooser2 = rankingLooser2 - rankingVariation;
  sr.getRange(rankingColumn+playersRankingRows[playerWinner1]).setValue(rankingWinner1); //Actualiza los resultados en la pestaña RANKING del excel
  sr.getRange(rankingColumn+playersRankingRows[playerWinner2]).setValue(rankingWinner2);
  sr.getRange(rankingColumn+playersRankingRows[playerLooser1]).setValue(rankingLooser1);
  sr.getRange(rankingColumn+playersRankingRows[playerLooser2]).setValue(rankingLooser2);
  return [rankingWinner1, rankingWinner2, rankingLooser1, rankingLooser2];
}

/**
 * playerWinner: prefijo de la pareja ganadora (Ejemplo Huertas-Garcho -> HU-GA)
 * playerLooser: prefijo de la pareja perdedora (Ejemplo Lax-Jorge -> LA-JO)
 */
function writePairRanking(playerWinner, playerLooser){
  var rankingWinner = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerWinner]).getValue()); //Obtiene el ranking actual de la pareja ganadora
  var rankingLooser = parseInt(sr.getRange(rankingColumn+playersRankingRows[playerLooser]).getValue()); //Obtiene el ranking actual de la pareja perdedora
  var rankingVariation = formuleRanking(rankingWinner, rankingLooser); //Calcula la variación del ranking
  rankingWinner = rankingWinner + rankingVariation; //suma a la pareja ganadora la variación del ranking
  rankingLooser = rankingLooser - rankingVariation; //resta a la pareja ganadora la variación del ranking
  sr.getRange(rankingColumn+playersRankingRows[playerWinner]).setValue(rankingWinner); //Actualiza los resultados en la pestaña RANKING del excel
  sr.getRange(rankingColumn+playersRankingRows[playerLooser]).setValue(rankingLooser);
  return [rankingWinner, rankingLooser];
}

function writeChartRankings(localPairArray, visitPairArray, rankingResultIndividuals, rankingResultPair, localWon){
  var i=0;

  //Buscamos la primera fila vacia y la inicializamos a 0 o con los resultados anteriores
  for(;i<500;i=i+1){
    if(rangeGraphics[i][0] === ""){
      rangeGraphics[i][0]=i+1; //Nº de partido
      for(var j=1;j<rangeGraphics[i].length;j=j+1){
        if(i===0){
          rangeGraphics[i][j]=0; //Si estamos al principio de la tabla de resultados ponemos un 0
        }else{
          rangeGraphics[i][j] = rangeGraphics[i-1][j]; //Si ya hay unos resultados contabilizados se copian los anteriores
        }

      }
      break;
    }
  }
  if(localWon){
    rangeGraphics[i][playersRankingGraphics[localPairArray[0]]]=rankingResultIndividuals[0]; //Actualizamos la tabla con el valor del ranking del jugador local 1 ganador
    rangeGraphics[i][playersRankingGraphics[localPairArray[1]]]=rankingResultIndividuals[1]; //Actualizamos la tabla con el valor del ranking del jugador local 2 ganador
    rangeGraphics[i][playersRankingGraphics[visitPairArray[0]]]=rankingResultIndividuals[2]; //Actualizamos la tabla con el valor del ranking del jugador visitante 1 perdedor
    rangeGraphics[i][playersRankingGraphics[visitPairArray[1]]]=rankingResultIndividuals[3]; //Actualizamos la tabla con el valor del ranking del jugador visitante 2 perdedor
    rangeGraphics[i][playersRankingGraphics[localPairArray[2]]]=rankingResultPair[0]; //Actualizamos la tabla con el valor de ranking de la pareja local ganadora
    rangeGraphics[i][playersRankingGraphics[visitPairArray[2]]]=rankingResultPair[1]; //Actualizamos la tabla con el valor de ranking de la pareja visitante perdedora

  }else{
    rangeGraphics[i][playersRankingGraphics[localPairArray[0]]]=rankingResultIndividuals[2]; //Actualizamos la tabla con el valor del ranking del jugador local 1 perdedor
    rangeGraphics[i][playersRankingGraphics[localPairArray[1]]]=rankingResultIndividuals[3]; //Actualizamos la tabla con el valor del ranking del jugador local 2 perdedor
    rangeGraphics[i][playersRankingGraphics[visitPairArray[0]]]=rankingResultIndividuals[0]; //Actualizamos la tabla con el valor del ranking del jugador visitante 1 ganador
    rangeGraphics[i][playersRankingGraphics[visitPairArray[1]]]=rankingResultIndividuals[1]; //Actualizamos la tabla con el valor del ranking del jugador visitante 2 ganador
    rangeGraphics[i][playersRankingGraphics[localPairArray[2]]]=rankingResultPair[1]; //Actualizamos la tabla con el valor de ranking de la pareja local perdedora
    rangeGraphics[i][playersRankingGraphics[visitPairArray[2]]]=rankingResultPair[0]; //Actualizamos la tabla con el valor de ranking de la pareja visitante ganadora
  }

}



/**
 * rankingWinner: es el valor actual del ranking del ganador
 * rankingLooser: es el valor actual del ranking del perdedor
 */

function formuleRanking(rankingWinner, rankingLooser){
  var delta = rankingWinner - rankingLooser; //Calcula la diferencia de rankings con respecto al ganador
  var rankingVariation = Math.round((-0.10)*delta + 20); //Aplica la formula para calcular la variación
  if(delta > 200){  //trunca la función, por encima
    return 1;
  }else if(delta < -200){ //trunca la función por debajo
    return 40;
  }else{
    return rankingVariation
  }
}

function showAlert(msg){
  SpreadsheetApp.getUi().alert(msg);
}




