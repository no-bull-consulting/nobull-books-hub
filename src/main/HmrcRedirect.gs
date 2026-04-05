/**
 * _hmrcRedirectPage(code)
 * Called from doGet when HMRC redirects back with ?code=XXX
 * Returns an HTML page that redirects the user back to the app with the code.
 */
function _hmrcRedirectPage(code) {
  var base = 'https://script.google.com/a/macros/nobull.consulting/s/';
  var dep  = 'AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j';
  var sid  = '1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg';
  var url  = base + dep + '/exec?id=' + sid + '&hmrc_code=' + encodeURIComponent(code);

  var lines = [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<meta charset="UTF-8">',
    '<meta http-equiv="refresh" content="1;url=' + url + '">',
    '<title>Connecting to HMRC</title>',
    '<style>',
    'body{font-family:sans-serif;background:#0f172a;color:#fff;',
    'display:flex;align-items:center;justify-content:center;',
    'min-height:100vh;margin:0;text-align:center}',
    'a{color:#60a5fa}',
    '</style>',
    '</head>',
    '<body>',
    '<div>',
    '<p>HMRC connection successful.</p>',
    '<p>Returning to no~bull books...</p>',
    '<p><a href="' + url + '">Click here if not redirected</a></p>',
    '</div>',
    '</body>',
    '</html>'
  ];

  return HtmlService.createHtmlOutput(lines.join(''))
    .setTitle('Connecting to HMRC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
