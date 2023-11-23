import g from '@battis/gas-lighter';

export function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Custom functions', 'popupDocumentation')
    .addToUi();
}

export function onInstall() {
  onOpen();
}

export function popupDocumentation() {
  openUrl(process.env.LIBRARY_URL);
}

/**
 * Open a URL in a new tab.
 * @see https://stackoverflow.com/a/47098533
 */
function openUrl(url) {
  var html = HtmlService.createHtmlOutput(
    '<html><script>' +
      'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};' +
      'var a = document.createElement("a"); a.href="' +
      url +
      '"; a.target="_blank";' +
      'if(document.createEvent){' +
      '  var event=document.createEvent("MouseEvents");' +
      '  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}' +
      '  event.initEvent("click",true,true); a.dispatchEvent(event);' +
      '}else{ a.click() }' +
      'close();' +
      '</script>' +
      // Offer URL as clickable link in case above code fails.
      '<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="' +
      url +
      '" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>' +
      '<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>' +
      '</html>'
  )
    .setWidth(90)
    .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening ...');
}

export function onHomepage() {
  return g.CardService.Card.create({
    name: 'Spreadsheet ICS',
    widgets: [
      JSON.stringify(
        PropertiesService.getUserProperties().getProperties(),
        null,
        2
      )
    ]
  });
}
