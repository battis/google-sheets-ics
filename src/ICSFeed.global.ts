function fold_(content) {
  const CRLF = '\r\n';
  return content
    .split('\n')
    .map((line) => {
      var folded = '';
      do {
        folded =
          folded +
          (folded.length > 0 ? CRLF + '\t' : '') +
          line.substring(0, 75);
        line = line.substring(75);
      } while (line.length > 0);
      return folded;
    })
    .join(CRLF);
}

export function doGet(e) {
  const match = e.parameter.feed.match(/(.+)\.(.+)/);
  const sheet = SpreadsheetApp.openById(match[1]);
  const a1Range = PropertiesService.getUserProperties().getProperty(
    `feed.${e.parameter.feed}`
  );
  const filename = e.parameter.filename || `${match[2]}.ics`;
  const content = sheet
    .getRange(a1Range)
    .getValues()
    .filter((val) => val.toString().length > 0)
    .join('\n');
  return ContentService.createTextOutput(fold_(content))
    .setMimeType(ContentService.MimeType.ICAL)
    .downloadAsFile(filename);
}
