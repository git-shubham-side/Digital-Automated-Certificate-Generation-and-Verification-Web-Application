const express = require("express");
const app = express();
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");
// body parser
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// static files
app.use(express.static(path.join(__dirname, "public")));

//Views
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

const upload = multer({ dest: "uploads/" });
// app.use();

// Routes
//Home Route
app.get("/", (req, res) => {
  res.render("../views/dashboard");
});

app.get("/generate-certificate", (req, res) => {
  res.render("../views/generate-certificate");
});

app.post("/generate-certificate", (req, res) => {
  const {
    studentName,
    rollNo,
    course,
    department,
    academicYear,
    semester,
    certificateType,
    eventName,
    eventDate,
    issuedBy,
  } = req.body;

  // for (e of data) {
  //   console.log(e);
  // }
  res.render("../views/certificate-prev", {
    studentName,
    rollNo,
    course,
    department,
    academicYear,
    semester,
    certificateType,
    eventName,
    eventDate,
    issuedBy,
  });
});

app.get("/login", (req, res) => {
  res.render("../views/login");
});
app.get("/register", (req, res) => {
  res.render("../views/college-regestration");
});

app.get("/certificate-type", (req, res) => {
  res.render("../views/certificate-type");
});
app.get("/preview-certificate", (req, res) => {
  res.render("../views/certificate-prev");
});

//Excel to JSON API
app.post("/excel-to-json", upload.single("file"), (req, res) => {
  const workbook = XLSX.readFile(req.file.path);
  const sheetName = workbook.SheetNames[0];
  const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  res.json({
    success: true,
    data: sheetData,
  });
  res.render;
});

// Server
app.listen(5000, () => console.log("Server is running"));
