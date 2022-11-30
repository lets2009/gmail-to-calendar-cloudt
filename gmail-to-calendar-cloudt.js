const QUERY_CLOUDT = 'subject:レッスン予約内容のご案内【クラウティ】 ';

function main() {
  pickUpMessage(QUERY_CLOUDT, function (message) {
    parseCloudt(message);
  });
}

function pickUpMessage(query, callback) {
  const messages = getMail(query);

  for (var i in messages) {
    for (var j in messages[i]) {
      const message = messages[i][j];
      // starは処理済みとする
      if (message.isStarred()) break;

      callback(message);

      message.star();
    }
  }
}

function getMail(query) {
  var threads = GmailApp.search(query, 0, 5);
  return GmailApp.getMessagesForThreads(threads);
}

function createEvent(title, description, location, year, month, dayOfMonth,
  startTimeHour, startTimeMinutes, endTimeHour, endTimeMinutes) {

  const calendar = CalendarApp.getDefaultCalendar();
  const startTime = new Date(year, month - 1, dayOfMonth, startTimeHour, startTimeMinutes, 0);
  const endTime = new Date(year, month - 1, dayOfMonth, endTimeHour, endTimeMinutes, 0);
  const option = {
    description: description,
    location: location,
  }

  console.log("start time: " + startTime);
  console.log("end time: " + endTime);
  calendar.createEvent(title, startTime, endTime, option);
}

// クラウティ
function parseCloudt(message) {
  const strDate = message.getDate();
  const strMessage = message.getPlainBody();

  const datePrefix = "レッスン時間：";
  const regexp = RegExp(datePrefix + '.*', 'gi');

  const result = strMessage.match(regexp);
  if (result == null) {
    console.log("This message doesn't have info.");
    return;
  }
  try {
    const parsedDate = result[0].replace(datePrefix, '');

    // 2022年12月1日 08:45〜08:55'

    const year = parsedDate.match(/[0-9]{4}年/gi)[0].replace('年', '');
    const month = parsedDate.match(/年[0-9]*月/gi)[0].replace('年', '').replace('月', '');
    const dayOfMonth = parsedDate.match(/月[0-9]*日/gi)[0].replace('月', '').replace('日', '');
    const startTimeHour = parsedDate.match(/[0-9]{2}:/gi)[0].replace(':', '');
    const startTimeMin = parsedDate.match(/:[0-9]{2}/gi)[0].replace(':', '');
    const endTimeHour = parsedDate.match(/[0-9]{2}:/gi)[1].replace(':', '');
    const endTimeMin = parsedDate.match(/:[0-9]{2}/gi)[1].replace(':', '');

    createEvent("英会話クラウティ", "mailDate: " + strDate,
      "", year, month, dayOfMonth, startTimeHour, startTimeMin, endTimeHour, endTimeMin);
  } catch(e) {
    console.error( e.stack);
    return;
  }
}

