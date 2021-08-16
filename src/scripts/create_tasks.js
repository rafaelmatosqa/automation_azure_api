const xlsx = require("xlsx");
const fs = require("fs");
const request = require("request");
const data = require("../../data.json");
const config = require("../config/config.json");

const wb = xlsx.readFile("./excelData.xlsx");
const ws = wb.Sheets["Planilha1"];
const rows = xlsx.utils.sheet_to_json(ws);
fs.writeFileSync("./data.json", JSON.stringify({ rows }));

let rowsLength;

rowsLength = rows.length;

for (let i = 0; i < rowsLength; i++) {
  request(
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json-patch+json",
      },
      auth: {
        username: config.username,
        password: config.password,
      },
      url: config.baseUrl + "/_apis/wit/workitems/$task?api-version=6.0",
      form: [
        {
          op: "add",
          path: "/fields/System.Title",
          from: null,
          value: data.rows[i].Title,
        },
        {
          op: "add",
          path: "/fields/System.Description",
          from: null,
          value: data.rows[i].Description,
        },
        {
          op: "add",
          path: "/fields/System.History",
          from: null,
          value: data.rows[i].Discussion,
        },
        {
          op: "add",
          path: "/fields/System.AssignedTo",
          from: null,
          value: data.rows[i].AssignedTo,
        },
        {
          op: "add",
          path: "/fields/System.State",
          from: null,
          value: data.rows[i].State,
        },
        {
          op: "add",
          path: "/fields/System.AreaPath",
          from: null,
          value: data.rows[i].AreaPath,
        },
        {
          op: "add",
          path: "/relations/-",
          value: {
            rel: "System.LinkTypes.Hierarchy-Reverse",
            url: data.rows[i].url_workitem,
            attributes: {
              comment: "Creating Parent Link",
            },
          },
        },
      ],
    },
    (error, response, body) => {
      console.log(body);
      console.log(response.body);
    }
  );
}
