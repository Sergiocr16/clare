if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js')
    .then(reg => console.log('SW registrado exitosamente ', reg))
    .catch(err => console.warn('Error al tratar de registrar SW ', err))
}
$(document).ready(function() {
  $('#my-file').val('')
    // $('select').material_select();
});

var fileInput  = document.querySelector( ".input-file" ),
    button     = document.querySelector( ".input-file-trig" ),
    the_return = document.querySelector(".file-return");

button.addEventListener( "keydown", function( event ) {
    if ( event.keyCode == 13 || event.keyCode == 32 ) {
        fileInput.focus();
    }
});
button.addEventListener( "click", function( event ) {
   fileInput.focus();
   return false;
});
function unfade(element) {
    var op = 0.1;  // initial opacity
    element.style.display = 'inline';
    var timer = setInterval(function () {
        if (op >= 1){
            clearInterval(timer);
        }
        element.style.opacity = op;
        element.style.filter = 'alpha(opacity=' + op * 100 + ")";
        op += op * 0.1;
    }, 10);
}
fileInput.addEventListener( "change", function( event ) {
    the_return.innerHTML = this.value;
    unfade(document.querySelector( "#process" ));
});
