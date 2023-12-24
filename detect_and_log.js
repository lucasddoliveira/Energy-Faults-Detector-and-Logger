function acharCidade() {
  var abaAtual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO');
  var sistemas = abaAtual.getRange("B2:B").getValues();

  sistemasOK = []
  for (let t = 0; t < sistemas.length; t++) {
    if(sistemas[t][0]!=''){

      sistemasOK.push(sistemas[t][0])
    }
  }
  //console.log(sistemasOK)
  cidades_paraiba = [
    "Água Branca",
    "Aguiar",
    "Alagoa Grande",
    "Alagoa Nova",
    "Alagoinha",
    "Alcantil",
    "Algodão de Jandaíra",
    "Alhandra",
    "Amparo",
    "Aparecida",
    "Araçagi",
    "Arara",
    "Araruna",
    "Areia",
    "Areia de Baraúnas",
    "Areial",
    "Aroeiras",
    "Assunção",
    "Baía da Traição",
    "Bananeiras",
    "Baraúna",
    "Barra de Santa Rosa",
    "Barra de Santana",
    "Barra de São Miguel",
    "Bayeux",
    "Belém",
    "Belém do Brejo do Cruz",
    "Bernardino Batista",
    "Boa Ventura",
    "Boa Vista",
    "Bom Jesus",
    "Bom Sucesso",
    "Bonito de Santa Fé",
    "Boqueirão",
    "Borborema",
    "Brejo do Cruz",
    "Brejo dos Santos",
    "Caaporã",
    "Cabaceiras",
    "Cabedelo",
    "Cachoeira dos Índios",
    "Cacimba de Areia",
    "Cacimba de Dentro",
    "Cacimbas",
    "Caiçara",
    "Cajazeiras",
    "Cajazeirinhas",
    "Caldas Brandão",
    "Camalaú",
    "Campina Grande",
    "Capim",
    "Caraúbas",
    "Carrapateira",
    "Casserengue",
    "Catingueira",
    "Catolé do Rocha",
    "Caturité",
    "Conceição",
    "Condado",
    "Conde",
    "Congo",
    "Coremas",
    "Coxixola",
    "Cruz do Espírito Santo",
    "Cubati",
    "Cuité",
    "Cuité de Mamanguape",
    "Cuitegi",
    "Curral de Cima",
    "Curral Velho",
    "Damião",
    "Desterro",
    "Diamante",
    "Dona Inês",
    "Duas Estradas",
    "Emas",
    "Esperança",
    "Fagundes",
    "Frei Martinho",
    "Gado Bravo",
    "Guarabira",
    "Gurinhém",
    "Gurjão",
    "Ibiara",
    "Igaracy",
    "Imaculada",
    "Ingá",
    "Itabaiana",
    "Itaporanga",
    "Itapororoca",
    "Itatuba",
    "Jacaraú",
    "Jericó",
    "João Pessoa",
    "Joca Claudino (ex-Santarém)",
    "Juarez Távora",
    "Juazeirinho",
    "Junco do Seridó",
    "Juripiranga",
    "Juru",
    "Lagoa",
    "Lagoa de Dentro",
    "Lagoa Seca",
    "Lastro",
    "Livramento",
    "Logradouro",
    "Lucena",
    "Mãe d'Água",
    "Malta",
    "Mamanguape",
    "Manaíra",
    "Marcação",
    "Mari",
    "Marizópolis",
    "Massaranduba",
    "Mataraca",
    "Matinhas",
    "Mato Grosso",
    "Matureia",
    "Mogeiro",
    "Montadas",
    "Monte Horebe",
    "Monteiro",
    "Mulungu",
    "Natuba",
    "Nazarezinho",
    "Nova Floresta",
    "Nova Olinda",
    "Nova Palmeira",
    "Olho d'Água",
    "Olivedos",
    "Ouro Velho",
    "Parari",
    "Passagem",
    "Patos",
    "Paulista",
    "Pedra Branca",
    "Pedra Lavrada",
    "Pedras de Fogo",
    "Pedro Régis",
    "Piancó",
    "Picuí",
    "Pilar",
    "Pilões",
    "Pilõezinhos",
    "Pirpirituba",
    "Pitimbu",
    "Pocinhos",
    "Poço Dantas",
    "Poço de José de Moura",
    "Pombal",
    "Prata",
    "Princesa Isabel",
    "Puxinanã",
    "Queimadas",
    "Quixaba",
    "Remígio",
    "Riachão",
    "Riachão do Bacamarte",
    "Riachão do Poço",
    "Riacho de Santo Antônio",
    "Riacho dos Cavalos",
    "Rio Tinto",
    "Salgadinho",
    "Salgado de São Félix",
    "Santa Cecília",
    "Santa Cruz",
    "Santa Helena",
    "Santa Inês",
    "Santa Luzia",
    "Santa Rita",
    "Santa Terezinha",
    "Santana de Mangueira",
    "Santana dos Garrotes",
    "Santo André",
    "São Bentinho",
    "São Bento",
    "São Domingos",
    "São Domingos do Cariri",
    "São Francisco",
    "São João do Cariri",
    "São João do Rio do Peixe",
    "São João do Tigre",
    "São José da Lagoa Tapada",
    "São José de Caiana",
    "São José de Espinharas",
    "São José de Piranhas",
    "São José de Princesa",
    "São José do Bonfim",
    "São José do Brejo do Cruz",
    "São José do Sabugi",
    "São José dos Cordeiros",
    "São José dos Ramos",
    "São Mamede",
    "São Miguel de Taipu",
    "São Sebastião de Lagoa de Roça",
    "São Sebastião do Umbuzeiro",
    "São Vicente do Seridó",
    "Sapé",
    "Serra Branca",
    "Serra da Raiz",
    "Serra Grande",
    "Serra Redonda",
    "Serraria",
    "Sertãozinho",
    "Sobrado",
    "Solânea",
    "Soledade",
    "Sossêgo",
    "Sousa",
    "Sumé",
    "Tacima",
    "Taperoá",
    "Tavares",
    "Teixeira",
    "Tenório",
    "Triunfo",
    "Uiraúna",
    "Umbuzeiro",
    "Várzea",
    "Vieirópolis",
    "Vista Serrana",
    "Zabelê"
  ]
  
  cidades = []
  
  for (let i = 0; i < sistemasOK.length; i++) {
    cidade = '---'

    switch (sistemasOK[i]){
      case 'SI-Acauã (A)':
        cidade = 'Itatuba'
        break;
      case 'SI-Tauá (A)':
        cidade = 'Araçagi'
        break;
      case 'SI-Canafístula II':
        cidade = 'Bananeiras'
        break;
      case 'SI-Canafístula II (A)':
        cidade = 'Bananeiras'
        break;
      case 'SI-Canafístula I (A)':
        cidade = 'Lagoa de Dentro'
        break;
      case 'SI-São Salvador':
        cidade = 'Caldas Brandão'
        break;
      case 'SI-São Salvador (A)':
        cidade = 'Caldas Brandão'
        break;
      case 'Mata Redonda':
        cidade = 'Alhandra'
        break;
      case 'SI-Capivara':
        cidade = 'Uiraúna'
        break;
      case 'SI-Capivara (A)':
        cidade = 'Uiraúna'
        break;
      case 'SI-Vaca Brava (A)':
        cidade = 'Remigio'
        break;
      case 'Jacumã (A)':
        cidade = 'Conde'
        break;
      case 'Olho dÁgua (A)':
        cidade = 'Olho dÁgua'
        break;
      case 'SI-Cariri (A)':
        cidade = 'Soledade'
        break;
      case 'SI-Coremas-Sabugi':
        cidade = 'Patos'
        break;
      default:
        cidade = '---'
        break;
    }

    if(cidade=='---'){
      //console.log('TESTE')
      for (let j = 0; j < cidades_paraiba.length; j++) {
        if (sistemasOK[i].includes(cidades_paraiba[j]) ) {
          cidade = cidades_paraiba[j]
        }
      }
    }

    cidades.push(cidade)
  }
  //console.log(cidades)
  for (var i = 1; i < cidades.length; i++) {
    abaAtual.getRange(i + 1, 17).setValue(cidades[i-1]);
  }


}

function acharUC(){
  var abaAtual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO');
  var abaAPOIO = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('APOIO');
  
  var sistemas = abaAtual.getRange("B2:B").getValues();
  var tags = abaAtual.getRange("D2:D").getValues();

  ucsOK = []

  // Valor que você deseja encontrar
  
  //console.log(linhaDestino[11])
  // A coluna onde você deseja procurar a correspondência
  var colunaProcurar = 2; // Coluna A, por exemplo
  var colunaSistema = 7;
  // A coluna onde você deseja obter a correspondência
  var colunaObter = 9; // Coluna B, por exemplo

  // Obtém os valores da coluna a ser procurada
  var valoresColunaProcurar = abaAPOIO.getRange(1, colunaProcurar, abaAPOIO.getLastRow(), 1).getValues();
  var valoresColunaProcurarSist = abaAPOIO.getRange(1, colunaSistema, abaAPOIO.getLastRow(), 1).getValues();
  
  //console.log(valoresColunaProcurarSist)
  
  for(var j =0;j<tags.length;j++){
    var correspondencia = '---'
    for (var i = 1; i < valoresColunaProcurar.length; i++) {
      
      
      if (valoresColunaProcurar[i][0] == tags[j]) {
        
        valorSistema = String(sistemas[j][0]).replace(' (A)', '');
        valorSistema = valorSistema.replace(' (E)', '');
        //console.log('---------------------')
        //console.log(valoresColunaProcurarSist[i][0])
        //console.log(valorSistema)
        
        if(valorSistema == valoresColunaProcurarSist[i][0]){
          //console.log('IGUAIS')
        // Correspondência encontrada, atribui o valor da coluna de obtenção à variável
          var correspondencia = abaAPOIO.getRange(i + 1, colunaObter).getValue();

          if(correspondencia==''){
            correspondencia = '---'
          }
          
          
          break;
        }
        //console.log('---------------------') // Pode interromper o loop se a correspondência for encontrada na primeira ocorrência
      }
      
    }
    ucsOK.push(correspondencia)
  }
  //console.log(ucsOK)

  for (var i = 1; i < ucsOK.length; i++) {
    abaAtual.getRange(i + 1, 19).setValue(ucsOK[i-1]);
  }

}

function acharInfluencia() {

  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaAPOIO = planilha.getSheetByName('APOIO2');
  var abaAtual = planilha.getSheetByName('HISTÓRICO-AUTOMÁTICO');
  var sistemas = abaAtual.getRange("B2:B").getValues();

  sistemasOK = []
  for (let i = 0; i < sistemas.length; i++) {
    if(sistemas[i][0]!=''){

      sistemasOK.push(sistemas[i][0])
    }
  }

  for (let x = 0; x < sistemasOK.length; x++) {
    sistemasOK[x] = sistemasOK[x].replace(' (A)', '');
    sistemasOK[x] = sistemasOK[x].replace(' (E)', '');
  }   
  //console.log(sistemasOK)
  // A coluna onde você deseja procurar a correspondência
  var colunaProcurar = 6; // Coluna A, por exemplo

  // A coluna onde você deseja obter a correspondência
  var colunaObter = 4; // Coluna B, por exemplo

  // Obtém os valores da coluna a ser procurada

  var valoresColunaProcurar = abaAPOIO.getRange(1, colunaProcurar, abaAPOIO.getLastRow(), 1).getValues();

  correspondenciasTotal = []

  for(var j = 0; j < sistemasOK.length; j++) {
    correspondencias = []
    for (var i = 1; i < valoresColunaProcurar.length; i++) {
      var correspondencia = '---'
      if (valoresColunaProcurar[i][0] == sistemasOK[j]) {
        
        // Correspondência encontrada, atribui o valor da coluna de obtenção à variável
        var correspondencia = abaAPOIO.getRange(i + 1, colunaObter).getValue();

        if(correspondencia==''){
          correspondencia = '---'
        }
        
        correspondencias.push(correspondencia) // Pode interromper o loop se a correspondência for encontrada na primeira ocorrência
      }
    }

    correspondenciasTotal.push(correspondencias)
  }
  //console.log(correspondenciasTotal)

  for(i=0;i<correspondenciasTotal.length;i++){
    correspondenciasTotal[i] = [...new Set(correspondenciasTotal[i])];
    correspondenciasTotal[i]= correspondenciasTotal[i].join(', ');
  }
  
  //console.log(correspondenciasTotal.length)
  for (var i = 1; i < correspondenciasTotal.length; i++) {
    correspondenciasTotal[i]= correspondenciasTotal[i].replace(', ---', '')
    abaAtual.getRange(i + 1, 18).setValue(correspondenciasTotal[i-1]);
  }
  
}

function isValidDate(date) {
  return !isNaN(date) && date instanceof Date;
}

function getSecondAppearance(str, targetChar) {
  let count = 0;

  for (let i = 0; i < str.length; i++) {
    if (str[i] === targetChar) {
      count++;
      if (count === 2) {
        return i; // Return the index of the second appearance
      }
    }
  }

  return -1; // If the character doesn't appear twice, return -1
}

function refreshDate() {
  var abaAtual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO');
  var valoresAbaAtual = abaAtual.getDataRange().getValues();
  
  for (var j = 1; j < valoresAbaAtual.length; j++) {
    
    var linhaAtual = valoresAbaAtual[j];
    
    var dataR = new Date(linhaAtual[7]);
    var data2 = new Date();
    
    if(linhaAtual[13]=='Não'){
      //console.log('AA')
      newDate = data2 - dataR

      var dias = Math.floor(newDate / (1000 * 60 * 60 * 24));
      newDate -= dias * 1000 * 60 * 60* 24;

      var horas = Math.floor(newDate / (1000 * 60 * 60));
      newDate -= horas * 1000 * 60 * 60;
    
      var minutos = Math.floor(newDate / (1000 * 60));
      newDate -= minutos * 1000 * 60;
    
      var segundos = Math.floor(newDate / 1000);

      // Formata o resultado no formato HH:MM:SS
      var novo = dias + 'D '+ horas + ':' + minutos + ':' + segundos;

      celula = abaAtual.getRange(j+1,11);

      celula.setValue(novo);
    } 
  }
}

function insertConcatenatedValues() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO'); // Replace 'Your_Sheet_Name' with your actual sheet name
  var lastRow = sheet.getLastRow();

  var columnCValues = sheet.getRange('C2:C' + lastRow).getValues(); // Get values from column C
  var columnDValues = sheet.getRange('D2:D' + lastRow).getValues(); // Get values from column D

  var concatenatedValues = [];

  for (var i = 0; i < columnCValues.length; i++) {
    var firstTwoChars = columnCValues[i][0].toString().substring(0, 2); // Get first two characters of column C
    var concatenated = firstTwoChars + '-' +columnDValues[i][0]; // Concatenate first two characters with content of column D
    concatenatedValues.push([concatenated]);
  }
  //concatenatedValues = concatenatedValues.flat()
  console.log(concatenatedValues)
  var targetRange = sheet.getRange('U2:U' + lastRow); // Set the range for column U
  targetRange.setValues(concatenatedValues); // Set the concatenated values in column U
}

function historico_ocorrencias() {
  // Link para a planilha de destino
  var linkPlanilhaDestino = 'https://docs.google.com/spreadsheets/d/1jUHl5SJmCRPl8HwCvo_6do6rHybW6HOl2rOwhAei6kc/edit#gid=0'; //10min

  //var linkPlanilhaDestino = 'https://docs.google.com/spreadsheets/d/1qODNmqpD07iR09dC-vbSw5uOIEpgjtC38Ozo200ve7A/edit#gid=1383778993' //45 dias

  // Condição que o item da coluna H deve atender
  var condicao = 'Falha de energia elétrica';

  // Abra a planilha de destino usando o link
  var planilhaDestino = SpreadsheetApp.openByUrl(linkPlanilhaDestino);

  // Acesse a aba da planilha de destino que você deseja comparar
  var abaDestino = planilhaDestino.getSheetByName('csv');

  // Acesse a aba da planilha atual que você deseja comparar
  var abaAtual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO');

  // Obtenha os valores das duas abas
  var valoresAbaDestino = abaDestino.getDataRange().getValues();

  var valoresAbaAtual = abaAtual.getDataRange().getValues();

  // Loop pelas linhas da aba de destino
  for (var f = 0; f < valoresAbaDestino.length; f++) {
    
    var linhaDestino = valoresAbaDestino[f];
    
    let concatenatedString = "";
    for (let i = 0; i < linhaDestino.length; i++) {
      concatenatedString += linhaDestino[i];
    }

    linhaDestino = concatenatedString.split('•')

    linhaDestino.unshift("", "");

    //console.log(linhaDestino)

    if (linhaDestino[7] === condicao) {
      
      regional = linhaDestino[3].substring(0, 2);
      
      switch (regional) {
        case 'ES':
          regional = 'ESPINHARAS'; 
          break;
        case 'LI':
          regional = 'LITORAL'; 
          break;
        case 'AP':
          regional = 'ALTO PIRANHAS'; 
          break;
        case 'BO':
          regional = 'BORBOREMA'; 
          break;
        case 'BR':
          regional = 'BREJO'; 
          break;
        case 'RP':
          regional = 'RIO DO PEIXE'; 
          break;
        default:
          regional = '-'; 
          break;
      }
      
      
      var data1 = new Date(linhaDestino[18]);
      var data2 = new Date(linhaDestino[15]);
      

      // Calcula a diferença em milissegundos
      var diferencaEmMilissegundos = data1 - data2;

      // Calcula a diferença em horas
      var dias = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60 * 24));
      diferencaEmMilissegundos -= dias * 1000 * 60 * 60* 24;

      var horas = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60));
      diferencaEmMilissegundos -= horas * 1000 * 60 * 60;
    
      var minutos = Math.floor(diferencaEmMilissegundos / (1000 * 60));
      diferencaEmMilissegundos -= minutos * 1000 * 60;
    
      var segundos = Math.floor(diferencaEmMilissegundos / 1000);

      // Formata o resultado no formato HH:MM:SS
      var resultadoFormatado = dias + 'D '+ horas + ':' + minutos + ':' + segundos;
      
      var resultadoFormatadoFIM = '---';
      
      var end = new Date(linhaDestino[31]);
      
      if(!isValidDate(end)){
        end = new Date(linhaDestino[32]);
      }
     
      if(linhaDestino[19]=='Não'){
        var now = new Date();

        var diferencaEmMilissegundos = now - data1;

        // Calcula a diferença em horas
        var dias = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60 * 24));
        diferencaEmMilissegundos -= dias * 1000 * 60 * 60* 24;

        var horas = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60));
        diferencaEmMilissegundos -= horas * 1000 * 60 * 60;
      
        var minutos = Math.floor(diferencaEmMilissegundos / (1000 * 60));
        diferencaEmMilissegundos -= minutos * 1000 * 60;
      
        var segundos = Math.floor(diferencaEmMilissegundos / 1000);

        // Formata o resultado no formato HH:MM:SS
        var resultadoFormatadoFIM = dias + 'D '+ horas + ':' + minutos + ':' + segundos;
      }else if(linhaDestino[19]=='Sim' && isValidDate(end)){
        
        var diferencaEmMilissegundos = end - data1;
        //console.log(data1);
        //console.log(end)
        // Calcula a diferença em horas
        var dias = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60 * 24));
        diferencaEmMilissegundos -= dias * 1000 * 60 * 60* 24;

        var horas = Math.floor(diferencaEmMilissegundos / (1000 * 60 * 60));
        diferencaEmMilissegundos -= horas * 1000 * 60 * 60;
      
        var minutos = Math.floor(diferencaEmMilissegundos / (1000 * 60));
        diferencaEmMilissegundos -= minutos * 1000 * 60;
      
        var segundos = Math.floor(diferencaEmMilissegundos / 1000);

        if(dias<0){
          dias = -dias
        }
        // Formata o resultado no formato HH:MM:SS
        var resultadoFormatadoFIM = dias + 'D '+ horas + ':' + minutos + ':' + segundos;
      }
      
      var planilha = SpreadsheetApp.getActiveSpreadsheet();
      
      var abaAPOIO = planilha.getSheetByName('APOIO');

      // Valor que você deseja encontrar
      correspondencia = '---'

      if(linhaDestino[35]==''){
        detalhamento = '---'
      }else{
        detalhamento = linhaDestino[35]
      }

      posicao = getSecondAppearance(linhaDestino[17],'-')
      var parte1 = linhaDestino[17].substring(0, posicao);
      var parte2 = linhaDestino[17].substring(posicao);

      coordenada = parte1 + "," + parte2
      dinheiro = '---'
      if(!isNaN(parseFloat(linhaDestino[34]))){
        dinheiro = parseFloat(linhaDestino[34])/100
      }

      protocolo = '---'
      
      if(linhaDestino[linhaDestino.length - 1]!='0'){   
        protocolo = linhaDestino[linhaDestino.length - 1]
      }

      

      linhaCerta = [linhaDestino[3],
            linhaDestino[9],
            regional,
            linhaDestino[11],
            linhaDestino[12],
            linhaDestino[5],
            linhaDestino[6],
            linhaDestino[18],
            linhaDestino[15],
            resultadoFormatado,
            resultadoFormatadoFIM,
            coordenada,
            detalhamento,
            linhaDestino[19],
            linhaDestino[20],
            dinheiro,
            '',
            '',
            correspondencia,
            protocolo
      ]
      //console.log(linhaDestino)
      // Verifique a condição na coluna H (índice 7)
      var abaAtual = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('HISTÓRICO-AUTOMÁTICO');
      var valoresAbaAtual = abaAtual.getDataRange().getValues();
      
      
      var correspondenciaEncontrada = false; 
    
      //console.log(linhaCerta)// Loop pelas linhas da aba atual
      for (var j = 1; j < valoresAbaAtual.length; j++) {
        
        
        var linhaAtual = valoresAbaAtual[j];
        //console.log(JSON.stringify(linhaCerta[0]))
        if(linhaCerta[0] == linhaAtual[0]){

          if(linhaAtual[13]=='Sim' && linhaCerta[13] == 'Não'){
            correspondenciaEncontrada = true;
          }
          if(linhaAtual[13]=='Não' && linhaCerta[13] == 'Sim'){
            correspondenciaEncontrada = true;
            x = j+1
            abaAtual.getRange('A' + x + ':T' + x).setValues([linhaCerta]); 
            //console.log('CHEGUEI')
          }
          if(linhaAtual[13]=='Não' && linhaCerta[13] == 'Não'){
            correspondenciaEncontrada = true;
          }
          if(linhaAtual[13]=='Sim' && linhaCerta[13] == 'Sim'){
            correspondenciaEncontrada = true;
          }
          //console.log('DUPLICADO')
          //correspondenciaEncontrada = true;

          

          break;
        }
        
      }
      //console.log(linhaDestino[18].toString())
      //console.log(linhaDestino)
      
      // Se não houver correspondência, adicione a linha à aba atual
      if (!correspondenciaEncontrada) {
        //console.log(linhaDestino[3])
        //console.log('AAA')
        abaAtual.appendRow(linhaCerta);
        //console.log(linhaDestino[38])
      }
    }
  }

  var dados = abaAtual.getDataRange()

  dados.sort({column: 8, ascending: false});

  refreshDate();

  var range = abaAtual.getRange("P2:P");
  var numberFormat = "R$ 0.00";
  var formatArray = [];

  for (var i = 1; i <= range.getNumRows(); i++) {
    formatArray.push([numberFormat]);
  }

  range.setNumberFormats(formatArray);

  acharCidade();

  acharInfluencia();

  acharUC();

  insertConcatenatedValues();

}
