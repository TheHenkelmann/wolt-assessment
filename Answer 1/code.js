const DEBUG = false;

const views = {
  variation:
    "Percentage of (max - min) / average for the observed period => identifies strong outliers. If combined with wow, it hints at a strong change",
  wow: "Percentage of average weekly change for the observed period => identifies consistent changes that are not outliers",
};
const viewsPerKpi = Object.keys(views).length;

const OPENAI_KEY = null;
const OPENAI_URL = "https://api.openai.com/v1/chat/completions";
const nameOfAnalysisSheet = "analysis_overview";

const subject = "Monthly KPI Report";
const urlOfGSheet =
  "https://docs.google.com/spreadsheets/d/17Aw5zkkl_6eDKp-SgrVgGiQVBFT7GiuSgFuvq-IJVyU/edit?usp=sharing";

// ----- sheet structure -----
const rowmap = {
  kpi: 0,
  direction: 1,
  view: 2,
};

const colmap = {
  area: 0,
};

const firstDataRow = Object.values(rowmap).reduce((a, b) => Math.max(a, b)) + 1;
const firstDataCol = Object.values(colmap).reduce((a, b) => Math.max(a, b)) + 1;

function sendEmail(tos, html) {
  // use the inbuilt MailApp service to send an email
  MailApp.sendEmail({
    to: tos.join(","),
    subject: subject,
    htmlBody: html,
  });
}

function parseData() {
  var ss =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameOfAnalysisSheet);
  let lastRow = ss.getLastRow();
  let lastCol = ss.getLastColumn();

  if (lastRow < firstDataRow || lastCol < firstDataCol) {
    throw new Error("No data found");
  } else if (lastRow == firstDataRow) {
    throw new Error("Only header found");
  } else if (lastCol == firstDataCol) {
    throw new Error("Only header found");
  } else if ((lastCol - firstDataCol) % viewsPerKpi != 0) {
    throw new Error("Invalid number of columns");
  }

  var data = ss.getRange(1, 1, lastRow, lastCol).getValues();
  var res = [];
  for (let row = firstDataRow; row < lastRow; row++) {
    for (let col = firstDataCol; col <= lastCol; col++) {
      let idxFirstOfSection =
        Math.floor((col - firstDataCol) / viewsPerKpi) * viewsPerKpi +
        firstDataCol;

      let kpi = data[rowmap.kpi][idxFirstOfSection];
      let direction = data[rowmap.direction][idxFirstOfSection];
      let view = data[rowmap.view][col];
      let value = data[row][col];

      let area = data[row][colmap.area];

      if (kpi && direction && view) {
        res.push({
          area: area,
          kpi: kpi,
          direction: direction,
          view: view,
          value: Math.round(value * 100) + "%",
        });
      }
    }
    // break;
  }

  // Logger.log(res);
  return res;
}

function getKPInames() {
  data = parseData();
  let res = [];
  data.forEach((d) => {
    let text = d.kpi + " (" + d.direction + ")";
    if (!res.includes(text)) {
      res.push(text);
    }
  });
  // Logger.log(res);
  return res;
}

function dataToCsv(data = []) {
  if (data.length == 0) {
    data = parseData();
  }
  let csv = "Area,KPI,Direction,View,Value\n";
  data.forEach((d) => {
    csv += `${d.area},${d.kpi},${d.direction},${d.view},${d.value}\n`;
  });
  return csv;
}

function callOpenAI(msg, systemMsg = null) {
  if (systemMsg == null) {
    let viewsMsg = Object.entries(views)
      .map(([k, v]) => `${k}: ${v}`)
      .join("\n");
    systemMsg = `You are a data analyst at a food delivery company. Summarize only the most important information. If there are any abnormalities, make sure to highlight them. If theres nothing to report, you can say that everything is normal.
I will give you the following KPIS: 
${getKPInames().join("\n")}. 
Each KPI has a direction which indicates whether a higher or lower value is better. 
For each KPI, I will give you ${viewsPerKpi} values: 
${viewsMsg}
`;
  }
  // Logger.log(systemMsg);

  let openaiBody = {
    model: "gpt-4o",
    temperature: 0.25,
    max_tokens: 2000,
    messages: [
      {
        role: "system",
        content: systemMsg,
      },
      {
        role: "user",
        content: msg,
      },
    ],
  };

  let options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + OPENAI_KEY,
    },
    payload: JSON.stringify(openaiBody),
  };

  Logger.log("------ OPENAI REQUEST ------");
  Logger.log(systemMsg);
  Logger.log(msg);
  if (DEBUG) {
    return "DEBUG\nOPENAI ANSWER";
  }

  let res = UrlFetchApp.fetch(OPENAI_URL, options);
  let resJson = JSON.parse(res);
  let analysis = resJson.choices[0].message.content;
  Logger.log(analysis);
  return analysis;
}

function writeAnalysis(data) {
  if (data == null) {
    data = parseData();
  }

  let msg = `Only report the most important information. Keep the message concise and to the point. Do not include all KPIs.
Answer only with a summary of the most important findings and as one message. If there's nothing to report, you can say that everything is normal. Structure your message as follows:
1. Name of area
2. Most severe problem if any
3. Most severe improvement if any
3. Severe problems / improvements

A development is severe if a wow change of +/- 10% happened, or a variation of > 50%
Severity ranks as follows:
1. nOrders, 
2. ADT Wolt
3. % Lateness > 25 min Wolt
4. Avg Delivery Rating Wolt


If you include a KPI, include the value and a very short analysis of the value.
Do not format the message. Include the name of the city you are analyzing in the message.
Here is the analysis of the KPIs for the past month in a CSV format:
${dataToCsv(data)}
`;
  // Logger.log(msg);

  return callOpenAI(msg);
}

function writeEmail() {
  let data = parseData();
  let areas = data.reduce((acc, d) => {
    if (!acc.includes(d.area)) {
      acc.push(d.area);
    }
    return acc;
  }, []);

  // areas = areas.slice(0, 1).concat(areas.slice(-2));
  Logger.log(`Writing email for areas: ${areas}`);

  var allAnylsis = [];
  for (let area of areas) {
    Logger.log(` > Anylyzing area: ${area}`);
    let subdata = data.filter((d) => d.area == area);
    allAnylsis.push(writeAnalysis(subdata));
  }
  Logger.log("Wrote analysis for all areas");

  // create a file from all analysis messages. format: area: analysis
  var text =
    "Detailed Analysis by Area\n-----\n\n" + allAnylsis.join("\n\n-----\n\n");
  let filename = new Date().toISOString().split("T")[0] + "_analysis.txt";
  let blob = Utilities.newBlob(text, "text/plain", filename);
  Logger.log("Wrote detailed analysis:");
  Logger.log(text);

  let msg = `I will provide you with a summary of the most important findings in general + a summary for each city. 
Write a concise text for the Head of Operations with the most important findings. Provide advice on the most urgent issues and key points he should be aware of. Keep in mind that the Head of Operations has a very tight schedule.

A development is severe if a wow change of +/- 10% happened, or a variation of > 50%
Severity ranks as follows:
1. nOrders,
2. ADT Wolt
3. % Lateness > 25 min Wolt
4. Avg Delivery Rating Wolt
5. Bundle Rate	
Alway strictly follow the severity ranking.


Structure your message as follows:
1. Talk about the general situation. You can include up to three cities with severe problems in the general analysis.
2. Then provide advice on the most urgent issues. Strictly follow the severity ranking. Regardless of how many cities have severe problems, only mention the most severe problem.
3. Then if necessary, mention specific cities. Strictly sort the cities by the severity ranking.

Do not talk about all KPIs, only the most important ones. Do not talk about all areas, only the most important ones.
Do not format your answer. Only use linebreaks where applicable. Keep it concise and to the point.


Example for a report:
1. nOrders (more is better)
GeneraL. small decrease of -1% wow
Cities: in Augsburg variation is at 136% and wow ist at 225% => consistent extraordinary growth

2. ADT Wolt (less is better)
General: small decrease of -2% wow
Cities: best wow: dortmund (-9%)

5. Bundle Rates: (more is better)
General: strong variation (63%) and consistend decrese by -11% wow
Cities: worst wow: hamburg (-25%), frankfurt (-24%), mannheim (-22%) | best wow: brunswick (+31%), leipzig (+30%), essen (+24%)


Here are the analyses for the past month:
${allAnylsis.join("\n\n")}
`;

  let generalAnalysis = callOpenAI(msg);
  Logger.log("Wrote general analysis:");
  Logger.log(generalAnalysis);

  html = HtmlService.createTemplateFromFile("mail");
  html.analysis = generalAnalysis;

  let message = html.evaluate().getContent();
  let to = "anton.henkelmann@gmail.com";

  if (!DEBUG || true) {
    let attachment = blob.getAs(MimeType.PLAIN_TEXT);
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: message,
      attachments: [attachment],
    });
  } else {
    Logger.log("DEBUG: Email not sent");
    Logger.log(message);
  }
}
