/*
 *  Libreria de Funciones JS para la interfaz de Calculo de Impuestos
 *      por : Manuel Alejandro Estevez Fernandez
 *
*/



$(document).ready(function(){
    var archivos_ ;
    $.fn.editable.defaults.mode = 'popup';
  /*
   *Este metodo esta ligado al evento onChange del FileLoader
   *en cuanto se detecte el archivo comenzara el upload
  */
  $('#mercancia_fileInput').on('change', preparaCarga);
    
    //function carga_guias(data) {
    //    /*
    //     *Recibe la informacion despues de los filtros de verificacion de solicitud exitosa
    //    */
    //    $('#tb_guias').empty();
    //    for (i =0; i<data.lentgh();i++){
    //        console.write(data[i]);
    //        //$('#tb_guias').append('<tr><td>'+data[i]['']+'</td><td>more data</td></tr>')
    //    }
    //};
    
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
                $('#tb_guias tbody').empty();
                $('#archivo_mercancia').removeClass('alert-success');
                break;
            case 'glp_fileInput':
                transmision_.append('tipo','guias');
                $('#div_archivo_glp').removeClass('alert-success');
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
                            switch(e.target.id){
                                case 'mercancia_fileInput':
                                    if (data.success == true) {
                                        $('#archivo_mercancia').addClass('alert-success');
                                        $('#tb_guias tbody').empty();
                                        for (var i =0; i<data.informacion.length;i++){
                                            if (data.informacion[i]['BL No.']!='BL No.') {
                                                if (data.informacion[i]['Repetida']==0) {
                                                    //console.log(data.informacion[i]);
                                                    $('#tb_guias').append('<tr><td>'+i+'</td><td>'+data.informacion[i]['BL No.']+'</td> <td><a href="#" id="fletes_'+i+'">0.00</a></td> <td><a href="#" id="seguros_'+i+'">0.00</a></td> <td><a href="#" id="embalajes_'+i+'" >0.00</a></td> <td><a href="#" id="otros_'+i+'" data-type="number"  [0-9]+([\.|,][0-9]+)? step="any" data-title="Ingresa los Otros" >0.00</a></td> <td> <a href="#" id="moneda_'+i+'"></a> </td> <td></td> </tr>');
                                                    $('#fletes_'+i).editable({
                                                        type : 'number',
                                                        step: 'any',
                                                        title: 'Captura el importe de Fletes',
                                                        });
                                                    $('#seguros_'+i).editable({
                                                        type : 'number',
                                                        step: 'any',
                                                        title: 'Captura el importe de Seguros',
                                                        });
                                                    $('#embalajes_'+i).editable({
                                                        type : 'number',
                                                        step: 'any',
                                                        title: 'Captura el importe de Embalajes',
                                                        });
                                                    $('#otros_'+i).editable({
                                                        type : 'number',
                                                        step: 'any',
                                                        title: 'Captura el importe de Otros Incrementables',
                                                        });
                                                    $('#moneda_'+i).editable({
                                                        type: 'select',
                                                        title: 'Selecciona la moneda',
                                                        placement: 'center',
                                                        value: 0,
                                                        source: [
                                                            {value:0, text:'Selecciona la Moneda'},
                                                            {value:1, text:'MXN'},
                                                            {value:2, text:'USD'},
                                                            {value:3, text:'EUR'},
                                                            {value:4, text:'KPW'},
                                                            {value:5, text:'KRW'},
                                                        ]
                                                    });
                                                }
                                            }
                                        };
                                        $('#div_archivo_glp').css('display','block');
                                    }else{
                                        alert(data.mensaje);
                                        e.target.focus();
                                    }
                                    
                                    break;
                                case 'glp_fileInput':
                                    $('#div_guiasfacturas').css('visible','display');
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