const _ = LodashGS.load();
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;

const HIKKI_URL = "HIKKI_URL";
const SLACK_POST_URL = "SLACK_POST_URL";
const CALENDAR_ID = "CALENDAR_ID";
const SECRET = "my_secret";
const EXCEPT_MEDIA = ["SPACE SHOWER TV Plus", "SPACE SHOWER TV", "MUSIC ON! TV", "MTV"];
const REMINDER_BEFORE_TIME = 15;
const PROP_NAMES_FOR_GSHEET = ["date", "startTime", "endTime", "media", "program"];

/**
 * Google Spread Sheetの場合、セルにデータを保存する際に勝手にデータの方を推測されて変な型で保存されるので、signatureをとっておく。
 * replaceしているのは、Spread Sheet内で[=, +]は式の一部とみなされてしまうため。
 *
 * @param {*} item
 * @returns
 */
const createSignature = (item: any) => {
  const message = [item.date, item.startTime, item.endTime, item.media, item.program].join();
  return Utilities.base64Encode(
    Utilities.computeHmacSignature(
      Utilities.MacAlgorithm.HMAC_SHA_512,
      message,
      SECRET,
    ),
  ).replace(/=/g, "E").replace(/\+/g, "P");
};

const postSlack = (hikkiItem: any) => {
  const message = "Hi @hasefumi23\n" +
    "I'll be on this TV program. Check it out!\n" +
    `_${hikkiItem.date} (${hikkiItem.startTime} - ${hikkiItem.endTime})_\n` +
    `*${hikkiItem.program}*\n` +
    `${hikkiItem.note.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g, "")}\n`;

  // TODO: ちゃんと画像とかを用意する
  const payload = {
    channel: "#hikki",
    icon_emoji: ":ghost:",
    text: message,
    username: "hikki info",
  };
  const options: URLFetchRequestOptions = {
    contentType: "application/json",
    method: "post",
    payload: JSON.stringify(payload),
  };

  UrlFetchApp.fetch(SLACK_POST_URL, options);
};

const createGcalendar = (item: any) => {
  const {
    title,
    media,
    startTime,
    endTime,
    description,
  } = item;
  const date = item.date.replace(/\./g, "/");

  // TODO: 予定にラベルを付けたい
  const event = CalendarApp.getCalendarById(CALENDAR_ID);
  const startAt = new Date(`${date} ${startTime}`);
  let endAt = null;
  if (_.isEmpty(endTime)) {
    // たまにendTimeに値が入ってこない場合があるので手当する
    endAt = new Date(date + " " + startTime);
    endAt.setHours(startAt.getHours() + 1);
  } else {
    endAt = new Date(date + " " + endTime);
  }

  event.createEvent(`[${media}]${title}`, startAt, endAt, { description });
  event.createEvent(`[${media}]${title}`, startAt, endAt, { description })
    .addPopupReminder(REMINDER_BEFORE_TIME);
};

const fetchUtadaJson = () => {
  const context = UrlFetchApp.fetch(HIKKI_URL).getContentText();
  // TODO: callbackをreplaceする処理を不要にできるはず
  const matched = context.match(/callback\(({.*})\)/) || "";
  return JSON.parse(matched[1].toString());
};

export default function main() {
  const logTime = new Date();
  Logger.log(`[${logTime.toString()}]START`);

  const utadaTvItems = fetchUtadaJson().items.tv.reverse();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TV_INFO_LIST");
  const lastRow = sheet.getLastRow();

  // データがない場合(初実行時のみの想定)は、データを突っ込んで処理を終了する
  if (lastRow === 0) {
    utadaTvItems.forEach((item: any) => {
      const signature = createSignature(item);
      const props = _.chain(item).pick(PROP_NAMES_FOR_GSHEET).values().value();
      sheet.appendRow([signature, ...props]);
    });

    return;
  }

  // 最新の行のsignatureを取得する
  const lastSignature = sheet.getSheetValues(lastRow, 1, 1, 4)[0][0];
  // fetchしてきたデータの配列でsignatureが一致するindexを探す
  const index = _.findIndex(utadaTvItems, (item: any) => {
    const signature = createSignature(item);
    return lastSignature === signature;
  });

  // まだ取り込んでいないデータに対してのみ各処理を実施する
  utadaTvItems.slice(index + 1).forEach((item: any) => {
    const signature = createSignature(item);
    const props = _.chain(item).pick(PROP_NAMES_FOR_GSHEET).values().value();
    sheet.appendRow([signature, ...props]);
    postSlack(item);

    if (!_.includes(EXCEPT_MEDIA, item.media)) {
      // 地デジで放送されているもののみカレンダーに登録したい
      createGcalendar(item);
    }
  });
}
