const Excel = require("exceljs");
const fs = require("fs");

(async () => {
  const workbook = new Excel.Workbook();
  const wb = await workbook.xlsx.readFile("./data.xlsx");

  const worksheet = wb.getWorksheet(1);
  const tests = [];
  const columns = {
    SUBJECT_ID: 1,
    GLUCOSE: 2,
    INSULIN: 3
  };

  worksheet.eachRow(row => {
    if (row.number !== 1) {
      const subjectMetadata = getSubjectIdMetadata(row.values[columns.SUBJECT_ID].toString());

      tests.push({
        ...subjectMetadata,
        subjectId: row.values[columns.SUBJECT_ID],
        glucose: row.values[columns.GLUCOSE],
        insulin: parseFloat(row.values[columns.INSULIN])
      });
    }
  });

  const dataWithExclusions = tests.filter(test => getLastNumberValue(test.preMRI) === '1');
  const consolidatedData = getConsolidatedData(dataWithExclusions);

  await createConsolidatedWorkbook(consolidatedData);

  // fs.writeFileSync(
  //   "./consolidated-data.json",
  //   JSON.stringify(consolidatedData, null, 2)
  // );
})();

function getLastNumberValue(num) {
  return num.toString().slice(-1);
}

function getSubjectIdMetadata(subjectId) {
    const regex = /1(\d{3})(\d{1})(\d{1})(\d{1})/;
    const result = regex.exec(subjectId);

    return {
      patientId: result[1],
      dayOfExperiment: result[2],
      preDiet: result[3],
      preMRI: result[4]
    };
}

function getConsolidatedData(dataset) {
  const consolidatedData = [];

  for (let i = 0; i < dataset.length; i++) {
    const current = dataset[i];
    const preDiet = {};
    const postDiet = {};

    let reading = {
      glucose: current.glucose,
      insulin: current.insulin
    };

    if (isPreDiet(current)) {
      preDiet[current.dayOfExperiment] = reading;
    } else {
      postDiet[current.dayOfExperiment] = reading;
    }

    for (let k = i+1; k < dataset.length; k++) {
      const next = dataset[k];

      if (next.patientId === current.patientId) {
        let reading = {
          glucose: next.glucose,
          insulin: next.insulin
        };

        if (isPreDiet(next)) {
          preDiet[next.dayOfExperiment] = reading;
        } else {
          postDiet[next.dayOfExperiment] = reading;
        }
      }
    }

    if (!consolidatedData.some(d => d.patientId === current.patientId)) {
      const pre = { 
        averageGlucose: getAverageGlucose(preDiet),
        averageInsulin: getAverageInsulin(preDiet)
      };

      const post = {
        averageGlucose: getAverageGlucose(postDiet),
        averageInsulin: getAverageInsulin(postDiet)
      };

      consolidatedData.push({
        patientId: current.patientId,
        preDiet: { ...preDiet, ...pre, HOMA: getHOMA(pre) },
        postDiet: { ...postDiet, ...post, HOMA: getHOMA(post) }
      });
    }
  }

  return consolidatedData;
}

function isPreDiet(patient) {
  return patient.preDiet === '1';
}

function getAverageGlucose(dataset) {
  let totalGlucose = 0;
  const daysWithGlucoseReadings = Object.keys(dataset)
    .filter(key => {
      const d = dataset[key];
      return d.glucose !== '' && d.glucose !== ' ' && d.glucose !== 0;
    });

  for (const day of daysWithGlucoseReadings) {
    totalGlucose += dataset[day].glucose;
  }

  return totalGlucose / daysWithGlucoseReadings.length;
}

function getAverageInsulin(dataset) {
  let totalInsulin = 0;
  const daysWithInsulinReadings = Object.keys(dataset)
    .filter(key => {
      const d = dataset[key];
      return d.insulin !== '' && d.insulin !== ' ' && d.insulin !== 0;
    });

  for (const day of daysWithInsulinReadings) {
    totalInsulin += dataset[day].insulin;
  }

  return totalInsulin / daysWithInsulinReadings.length;
}

function getHOMA(data) {
  return (data.averageGlucose * data.averageInsulin) / 405;
}

function createConsolidatedWorkbook(consolidatedData) {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Consolidated Data');

  worksheet.addRow(['patient_id', 'average_glucose_pre', 'average_insulin_pre', 'homa_pre', 'average_glucose_post', 'average_insulin_post', 'homa_post']);

  for (const patient of consolidatedData) {
    worksheet.addRow([
      patient.patientId,
      patient.preDiet.averageGlucose,
      patient.preDiet.averageInsulin,
      patient.preDiet.HOMA,
      patient.postDiet.averageGlucose,
      patient.postDiet.averageInsulin,
      patient.postDiet.HOMA,
    ]);
  }

  return workbook.xlsx.writeFile('./consolidated-data.xlsx');
}