const xlsxFile = require("read-excel-file/node");
const fs = require("fs");

const EXCEL_FILES = [
  {
    fileName: "2022-02",
    indexArray: [5, 2, 6, 8, 10, 11],
  },
  {
    fileName: "2022-01",
    indexArray: [5, 2, 6, 8, 10, 13],
  },
];

(async function () {
  for (const { fileName, indexArray } of EXCEL_FILES) {
    await convertExcelToJson(fileName, indexArray);
  }
})();

async function convertExcelToJson(fileName, indexArray) {
  const rows = await xlsxFile(`./${fileName}.xlsx`);
  const json = [];

  for (const row of rows) {
    let index = 0;
    const classNumber = row[indexArray[index++]];
    const category = row[indexArray[index++]];
    const name = row[indexArray[index++]];
    const professor = row[indexArray[index++]];
    const rawSchedules = row[indexArray[index++]];
    const processedSchedules = processRawSchedules(rawSchedules);
    const isELerning = processedSchedules === null;
    const grades = row[indexArray[index++]];

    json.push({
      classNumber,
      category,
      name,
      professor,
      schedules: processedSchedules,
      isELerning,
      grades,
    });
  }
  fs.writeFileSync(`./${fileName}.json`, JSON.stringify(json));
}

function processRawSchedules(rawSchedules) {
  if (!rawSchedules) return null;

  const processedSchedules = [];

  const lectureRoomChunks = rawSchedules.split(" ");
  for (const lectureRoomChunk of lectureRoomChunks) {
    const rawSchedulesString = lectureRoomChunk.replace("]", "").split(":")[1];
    const scheduleStrings = rawSchedulesString.split(",");

    const scheduleObjects = convertScheduleStringsToObjects(scheduleStrings);
    const convertedSchedules = convertSchedulesTimeUnit(scheduleObjects);

    // 한 강의실에서 2교시 이상 수업할 때가 있고
    // 두 강의실에서 1교시 씩 수업할 때가 있다.
    processedSchedules.push(...convertedSchedules);
  }

  return processedSchedules;
}

function convertScheduleStringsToObjects(scheduleStrings) {
  const scheduleObjects = [];
  for (const scheduleString of scheduleStrings) {
    const [day, ...times] = scheduleString.replace(/\)/g, "").split("(");
    const startTimeString = times[0];
    const endTimeString = times[times.length - 1];

    scheduleObjects.push({ day, startTimeString, endTimeString });
  }
  return scheduleObjects;
}

function convertSchedulesTimeUnit(scheduleObjects) {
  const convertedSchedules = [];
  for (const { day, startTimeString, endTimeString } of scheduleObjects) {
    const { startHour, startMinute } = TIME_UNITS[startTimeString];
    const { endHour, endMinute } = TIME_UNITS[endTimeString];

    const startMinutes = startHour * 60 + startMinute;
    const endMinutes = endHour * 60 + endMinute;
    const workingMinutes = endMinutes - startMinutes;

    convertedSchedules.push({ day, startHour, startMinute, workingMinutes });
  }
  return convertedSchedules;
}

const TIME_UNITS = {
  0: { startHour: 8, startMinute: 0, endHour: 8, endMinute: 50 },
  1: { startHour: 9, startMinute: 0, endHour: 9, endMinute: 50 },
  2: { startHour: 10, startMinute: 0, endHour: 10, endMinute: 50 },
  3: { startHour: 11, startMinute: 0, endHour: 11, endMinute: 50 },
  4: { startHour: 12, startMinute: 0, endHour: 12, endMinute: 50 },
  5: { startHour: 13, startMinute: 0, endHour: 13, endMinute: 50 },
  6: { startHour: 14, startMinute: 0, endHour: 14, endMinute: 50 },
  7: { startHour: 15, startMinute: 0, endHour: 15, endMinute: 50 },
  8: { startHour: 16, startMinute: 0, endHour: 16, endMinute: 50 },
  9: { startHour: 17, startMinute: 0, endHour: 17, endMinute: 50 },

  "0A-0": { startHour: 7, startMinute: 30, endHour: 8, endMinute: 45 },
  "1-2A": { startHour: 9, startMinute: 0, endHour: 10, endMinute: 15 },
  "2B-3": { startHour: 10, startMinute: 30, endHour: 11, endMinute: 45 },
  "4-5A": { startHour: 12, startMinute: 0, endHour: 13, endMinute: 15 },
  "5B-6": { startHour: 13, startMinute: 30, endHour: 14, endMinute: 45 },
  "7-8A": { startHour: 15, startMinute: 0, endHour: 16, endMinute: 15 },
  "8B-9": { startHour: 16, startMinute: 30, endHour: 17, endMinute: 45 },

  야1: { startHour: 18, startMinute: 0, endHour: 18, endMinute: 50 },
  야2: { startHour: 18, startMinute: 55, endHour: 19, endMinute: 45 },
  야3: { startHour: 19, startMinute: 50, endHour: 20, endMinute: 40 },
  야4: { startHour: 20, startMinute: 45, endHour: 21, endMinute: 35 },
  야5: { startHour: 21, startMinute: 40, endHour: 22, endMinute: 30 },
  야6: { startHour: 22, startMinute: 35, endHour: 23, endMinute: 25 },

  "야1-2A": { startHour: 18, startMinute: 0, endHour: 19, endMinute: 15 },
  "야2B-3": { startHour: 19, startMinute: 25, endHour: 20, endMinute: 40 },
  "야4-5A": { startHour: 20, startMinute: 50, endHour: 22, endMinute: 5 },
};
