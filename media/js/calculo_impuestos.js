/*
 *  Libreria de Funciones JS para la interfaz de Calculo de Impuestos
 *      por : Manuel Alejandro Estevez Fernandez
 *
*/



$(document).ready(function(){
    var archivos_ ;
  /*
   *Este metodo esta ligado al evento onChange del FileLoader
   *en cuanto se detecte el archivo comenzara el upload
  */
  $('#mercancia_fileInput').on('change', preparaCarga);
  
    
    function preparaCarga(e) {
        archivos_ = e.target.files;
        subeArchivos(e);
    };
    
    function subeArchivos(e){
        // Evitando que se envien eventos de mas
        e.stopPropagation();
        e.preventDefault();
        // Preparando la informacion a enviar
        var transmision_ = new FormData();
        // SE asigna el tipo dependiendo el input cargado.
        switch (e.target.id) {
            case 'mercancia_fileInput' :
                transmision_.append('tipo','mercancia'); 
                break;
            case 'glp_fileInput':
                transmision_.append('tipo','guias');
                break;
            default:
                transmision_.append('tipo','vacio');
        };
        
        $.each(archivos_, function(key,value){
            transmision_.append(key,value);
        });
        // Realizando la consulta
        $.ajax({
            url:  $('#archivo_mercancia').attr('action'),
            method: $('#archivo_mercancia').attr('method'),
            data: transmision_,
            dataType: 'json',
            processData: false,
            contentType: false,
            success: function(data, textStatus, jqXHR){
                         if(typeof data.error === 'undefined')
                         {
                            // Success so call function to process the form
                            alert(data['correcto'])
                            if (e.target.id == 'mercancia_fileInput') {
                                $('#div_archivo_glp').style
                            }
                         }
                         else
                         {
                             // Handle errors here
                             console.log('ERRORES: ' + data.error);
                         }
                     },
            error: function(jqXHR, textStatus, errorThrown){
                        // Handle errors here
                        console.log('ERRORS: ' + textStatus);
                        // STOP LOADING SPINNER
                    }
           
        });
    }
});