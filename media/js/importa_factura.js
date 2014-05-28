/*  Libreria de funciones JS para la seccion 
 *      Importa Facturas
 *
 *  por: Manuel Alejandro Estevez Fernandez
 *      Abril 2014
*/

function carga_archivo(e) {
    //code
    e.preventDefault();
    $('#div_archivo').css('display','block');
}
/*
function upload(e) {
    e.preventDefault();
    $.ajax({
        url: $(this).attr('action'),
        type: $(this).attr('method'),
        cache: false,
        data: $(this).serialize(),
        files ,
        async: false ,
        success: function(data) {
            $('#mercancias_factura').empty();
            alert(data);
            //$('#detalle_mercancias_factura').append('<tr><td>my data</td><td>more data</td></tr>')
        },
        error: function(data) {
            
        },
        
    });
    //return false;
}*/