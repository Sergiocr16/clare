(function($){
  $(function(){

    $('.button-collapse').sidenav();
    $('.parallax').parallax();
    $('input#input_text, textarea#textarea1').characterCounter();
    $('#textarea1').val('New Text');
    $('#textarea1').trigger('autoresize');
    $('.modal').modal();
    $('#modal1').modal('open');
    $('.slider').slider();
    $('.collapsible').collapsible();
    $('ul.tabs').tabs();
    // $('ul.tabs').tabs('select_tab', 'tab_id');
    $('.materialboxed').materialbox();

    // $('.datepicker').pickadate({
    // selectMonths: true, // Creates a dropdown to control month
    // selectYears: 50, // Creates a dropdown of 15 years to control year,
    // today: 'Today',
    // clear: 'Clear',
    // close: 'Ok',
    // closeOnSelect: false // Close upon selecting a date,
    // });

    });
})(jQuery);
