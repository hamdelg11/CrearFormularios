function CrearFormulario() {

  var sheet = SpreadsheetApp.openById("1AxAgmb8dTDdpb9-25dO7WowQsE_dXWGVhQocXlcNx4k").getSheetByName('Hoja 1');
  var lastRow = sheet.getLastRow(); //Obtiene la última fila
  var lastCol = sheet.getLastColumn(); //Obtiene la última columna

  var tituloForm= sheet.getRange(1,1).getValues()[0]; //Obtiene el valor del titulo de la primera celda
  var data = sheet.getDataRange().getValues(); //Rango de la información en el sheets

  var form = FormApp.create(tituloForm)  //Creación del formulario
       .setTitle(tituloForm);  


  data.forEach(function(item){ //Recorrer todas las filas
      item.forEach(function(elemento){  //Recorrer cada elemento del arreglo de las filas
    
        switch(elemento){
            case 'TEXTO':
                var requerido= true;

                if(item[2]=='no'){
                  requerido = false;
                }
                form.addTextItem()  
                .setTitle(item[0])
                .setRequired(requerido);
                break;

            case 'MULTIPLE':
                var requerido= true;
                var opciones = item.slice(3,lastCol) //Obtiene todas las respuestas para la pregunta
                var filtrado = opciones.filter(opcion => opcion != '') //Elimina celdas vacias
                if(item[2]=='no'){ 
                  requerido = false;
                }
                form.addMultipleChoiceItem()  
                .setTitle(item[0])
                .setChoiceValues(filtrado)
                .setRequired(requerido);
                break;        
            case 'CHECKBOX':
                var requerido= true;
                var opciones = item.slice(3,lastCol) //Obtiene todas las respuestas para la pregunta
                var filtrado = opciones.filter(opcion => opcion != '') //Elimina celdas vacias
                if(item[2]=='no'){
                  requerido = false;
                }
                
                form.addCheckboxItem()  
                .setTitle(item[0]) 
                .setChoiceValues(filtrado)
                .setRequired(requerido);
                break;

            case 'LISTA': 
                var requerido= true;
                var opciones = item.slice(3,lastCol) //Obtiene todas las respuestas para la pregunta
                var filtrado = opciones.filter(opcion => opcion != '') //Elimina celdas vacias
                if(item[2]=='no'){
                  requerido = false;
                }
                
                form.addListItem()  
                .setTitle(item[0]) 
                .setChoiceValues(filtrado)
                .setRequired(requerido);
                break;   
          }
        
    })
  })
    



    var url = form.getPublishedUrl(); //Obtener URL del formulario


    Logger.log(url);
}

