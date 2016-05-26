function onOpen() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('url shortener', 'start')
  .addToUi();
}

function start(){
  var userInterface = HtmlService.createTemplateFromFile('sidebar').evaluate();
  SpreadsheetApp.getUi().showSidebar(userInterface);
}

function createReport(from, count, isCurrentSheet){
  count = isNumeric(count) ? count : 0;
  var d = new Date();
  var startToken = Utilities.formatString('%s-%02d-%02dT%02d:%02d:%02d.000+00:00', d.getYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(),  d.getSeconds());
  if(from){
    var d = from.split('-');
    if(d.length === 3) {
      startToken = Utilities.formatString('%s-%s-%sT00:00:00.000+00:00', d[0], d[1], d[2]);
    }
  }
  getData(startToken, count, isCurrentSheet);
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function getData_(){
  var ssh = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ssh.insertSheet('' + (new Date().getTime()));
  var values = UrlShortener.Url.list({projection: 'FULL'}).items.map(function(item){
    return [item.id, item.longUrl, item.analytics.allTime.shortUrlClicks, item.analytics.day.shortUrlClicks, item.analytics.week.shortUrlClicks, item.analytics.month.shortUrlClicks];
  });
  //Logger.log(values);
  values.unshift(['id', 'longUrl', 'allTime', 'day', 'week', 'month']);
  sh.getRange(1, 1, values.length, values[0].length).setValues(values);
  sh.setFrozenRows(1);
  sh.activate();                       
}

function getData2(startToken){
  var optionalArgs = {projection: 'FULL'};
  startToken && (optionalArgs['start-token'] = startToken);
  return UrlShortener.Url.list(optionalArgs);
}

function getData(startToken, count, isCurrentSheet){
  startToken = startToken || undefined;
  if(count < 1 || count > 150) count = 150;
  var ssh = SpreadsheetApp.getActiveSpreadsheet();
  var sh = isCurrentSheet ? ssh.getActiveSheet() : ssh.insertSheet('' + (new Date().getTime()));
  var list = getData2(startToken);
  var values = [];
  while(values.length < count){
    var part = list.items.map(function(item){
      return [startToken, item.id, item.longUrl, item.analytics.allTime.shortUrlClicks, item.analytics.day.shortUrlClicks, item.analytics.week.shortUrlClicks, item.analytics.month.shortUrlClicks];
    });
    values = values.concat(part);
    if(!list.nextPageToken) break;
    startToken = list.nextPageToken;
    list = getData2(startToken);    
  }
  values.unshift(['page', 'id', 'longUrl', 'allTime', 'day', 'week', 'month']);
  sh.clear().getRange(1, 1, values.length, values[0].length).setValues(values);
  sh.setFrozenRows(1);
  sh.activate();     
}

function printData(){
  var d = new Date();
  var startToken = Utilities.formatString('%s-%02d-%02dT%02d:%02d:%02d.000+00:00', d.getYear(), d.getMonth() + 1, d.getDate(), d.getHours(), d.getMinutes(),  d.getSeconds());
  Logger.log(startToken);
}
