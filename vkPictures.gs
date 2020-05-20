function vkPictures() {

  var k=50;      //ограничение по количеству запросов (в одном запросе 100 постов)
  
  var masPics = [["Id поста", "дата", "Подпись к картинке", "Картинка из поста (не ссылка)"]];
  
  var owner_id = 31480508;       //id сообщества
  var access_token = "";
  var v = 5.103;
  
  for (n=0; n<=k*100; n+=100) {
    
    var response = UrlFetchApp.fetch("https://api.vk.com/method/wall.get?owner_id=-"+owner_id+"&offset="+n+"&count=100&access_token="+access_token+"&v="+v);

    posts = JSON.parse(response.getContentText());
    
    if (n>posts.response.count) break;
    
    for (let item of posts.response.items) {
      if (item.hasOwnProperty('attachments')) {
        for (let attachment of item.attachments) {
          if (attachment.type==="photo") {
            if (attachment.photo!=undefined) {
              let date = new Date(item.date*1000)
              masPics.push([item.id, date, item.text, "=IMAGE(\""+attachment.photo.sizes[4].url+"\")"]);
            }
          }
        }
      }
    }
    
  }
  
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  
  sh.clear();
  
  var len = masPics.length;
  
  sh.getRange(1, 1, len, 4).setValues(masPics);
  sh.getRange(1, 2, len, 2).setNumberFormat('dd.MM.yyyy H:mm:ss');
  sh.setRowHeights(2, len-1, 159);
}

//автообновление при открытии
function onOpen(){
  vkPictures();
}
