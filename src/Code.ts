// TODO: gitで管理する

const _ = LodashGS.load();
// import _ from "lodash";
import URLFetchRequestOptions = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions;
import settings from "./settings.json";

const HIKKI_URL = settings.HIKKI_URL;
const SLACK_POST_URL = settings.SLACK_POST_URL;
const SECRET = "my_secret";
const EXCEPT_MEDIA = ["SPACE SHOWER TV Plus", "SPACE SHOWER TV", "MUSIC ON! TV", "MTV"];
const REMINDER_BEFORE_TIME = 15;

/**
 * Google Spread Sheetの場合、セルにデータを保存する際に勝手にデータの方を推測されて変な型で保存されるので、signatureをとっておく。
 * replaceしているのは、Spread Sheet内で[=, +]は式の一部とみなされてしまうため。
 *
 * @param {*} item
 * @returns
 */
function createSignature(item: any) {
  const message = [item.date, item.startTime, item.endTime, item.media, item.program].join();
  return Utilities.base64Encode(
    Utilities.computeHmacSignature(
      Utilities.MacAlgorithm.HMAC_SHA_512,
      message,
      SECRET,
    ),
  ).replace(/=/g, "E").replace(/\+/g, "P");
}

function postSlack(message: any) {
  // TODO: ちゃんと画像とかを用意する
  // TODO: htmlタグや実体参照などのサニタイズする
  // TODO: slackへの投稿をもうちょっとおしゃれにする
  // TODO: 自分にメンションが飛ぶようにする
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
}

function createGcalendar(title: string, media: string, date: string, startTime: string, endTime: string, description: string) {
  // TODO: 予定にラベルを付けたい
  const event = CalendarApp.getCalendarById("eqkvr0mmooe0l3c2mo7ailq0ag@group.calendar.google.com");
  const startAt = new Date(date + " " + startTime);
  let endAt = null;
  if (_.isEmpty(endTime)) {
    // たまにendTimeに値が入ってこない場合があるので手当する
    endAt = new Date(date + " " + startTime);
    endAt.setHours(startAt.getHours() + 1);
  } else {
    endAt = new Date(date + " " + endTime);
  }

  event.createEvent("[" + media + "]" + title, startAt, endAt, { description })
    .addPopupReminder(REMINDER_BEFORE_TIME);
}

function fetchUtadaJson() {
  const context = UrlFetchApp.fetch(HIKKI_URL).getContentText();
  // TODO: callbackをreplaceする処理を不要にできるはず
  const matched = context.match(/callback\(({.*})\)/) || "";
  const json = JSON.parse(matched[1].toString());

  return json;
}

export default function main() {
  const logTime = new Date();
  Logger.log("[" + logTime.toString() + "]" + "START");

  const utadaJson = fetchUtadaJson();
  const utadaTvItems = utadaJson.items.tv.reverse();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TV_INFO_LIST");
  const lastRow = sheet.getLastRow();

  // データがない場合(初実行時のみの想定)は、データを突っ込んで処理を終了する
  if (lastRow === 0) {
    utadaTvItems.forEach((item: any) => {
      const signature = createSignature(item);
      sheet.appendRow([signature, item.date, item.startTime, item.endTime, item.media, item.program]);
    });

    return;
  }

  // 最新の行のsignatureを取得する
  const lastSignature = sheet.getSheetValues(lastRow, 1, 1, 4)[0][0];

  // fetchしてきたデータの配列でsignatureが一致するindexを探す
  const index = _.findIndex(utadaTvItems, (item) => {
    const signature = createSignature(item);
    return lastSignature === signature;
  });

  // まだ取り込んでいないデータに対してのみ各処理を実施する
  utadaTvItems.slice(index + 1).forEach((item: any) => {
    const signature = createSignature(item);

    sheet.appendRow([signature, item.date, item.startTime, item.endTime, item.media, item.program]);
    postSlack("@hasefumi23 " + "\n"
      + "_" + item.date + " " + "(" + item.startTime + "-" + item.endTime + ")" + "_"
      + "\n" + "*" + item.program + "*" + "\n"
      + _.repeat("=", 90)
      + "\n" + item.note + "\n"
      + _.repeat("=", 90),
    );
    if (!_.includes(EXCEPT_MEDIA, item.media)) {
      // 地デジで放送されているもののみカレンダーに登録したい
      createGcalendar(item.program, item.media, item.date.replace(/\./g, "/"), item.startTime, item.endTime, item.note);
    }
  });
}
