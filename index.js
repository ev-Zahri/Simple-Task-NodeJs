const express = require("express");
const requestIp = require("request-ip");
const si = require('systeminformation');
const os = require("os");
const ExcelDataManager = require("./logic");

const app = express();
const port = 3000;

app.use(requestIp.mw());
app.get("/", (req, res) => {
  const excelDataManager = new ExcelDataManager();
  const date = new Date();
  const currentTime =
    `${date.getHours()}-${date.getMinutes()}`;
  let currentDate =
    `${date.getDate()}-${date.getMonth() + 1}-${date.getFullYear()}`;

  const cpuInfo = os.cpus()[0];

  const clientInfo = {
    // typeDevice: navigator.userAgent,
    deviceName: os.hostname(),
    modelName: os.userInfo().username, // Change this to the appropriate field
    ipAddress: req.clientIp,
    architecture: os.arch(),
    processor: `${cpuInfo.model} @ ${cpuInfo.speed} MHz`,
    currentTime,
    currentDate,
  };
  excelDataManager.saveDataClient(clientInfo);
  res.send("Data saved");
});

app.listen(port, () => {
  console.log(`Website berjalan pada port ${port}`);
});
