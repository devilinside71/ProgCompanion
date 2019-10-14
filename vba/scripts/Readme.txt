felbontás elemekre
declaration
Dim xxx As String|Double


Formatter:
line break 80 char
1., REM line
(\S+)


vba.json snippets

https://api.jquery.com/jQuery.getJSON/
https://codepen.io/KryptoniteDove/post/load-json-file-locally-using-pure-javascript

$.getJSON( "ajax/test.json", function( data ) {
  var items = [];
  $.each( data, function( key, val ) {
    items.push( "<li id='" + key + "'>" + val + "</li>" );
  });
 
  $( "<ul/>", {
    "class": "my-new-list",
    html: items.join( "" )
  }).appendTo( "body" );
});

