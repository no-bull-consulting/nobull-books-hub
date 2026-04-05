function _hmrcRedirectPage(code) {
  var base = 'https://script.google.com/a/macros/nobull.consulting/s/';
  var dep  = 'AKfycbxAr1fwnaEmr5Q3tD8_hOrj8zsQ8TtcAofQipYASdEDR4tKJG8liN-OEMIL1nnrka5j';
  var sid  = '1gIFwQUtbhGaM3HIHbFFaT7lIAU4BN3IksAOv1_uuUKg';
  var url  = base + dep + '/exec?id=' + sid + '&hmrc_code=' + encodeURIComponent(code);
  var tmpl = HtmlService.createTemplateFromFile('HmrcRedirect');
  tmpl.redirectUrl = url;
  return tmpl.evaluate()
    .setTitle('Connecting to HMRC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
