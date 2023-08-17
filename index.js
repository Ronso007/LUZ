var createError = require("http-errors");
var express = require("express");
var path = require("path");
var cookieParser = require("cookie-parser");
var logger = require("morgan");
var cors = require("cors");
const { google } = require("googleapis");
require("dotenv").config();
var fs = require("fs");

var mainRoute = require("./routes/mainRoute");

var app = express();
const oauth2Client = new google.auth.OAuth2(process.env.CLIENT_ID, process.env.CLIENT_SECRET, process.env.REDIRECT_URI);
oauth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });
global.oauth2Client = oauth2Client;

// try {
//   const creds = fs.readFileSync("creds.json");
//   oauth2Client.setCredentials(JSON.parse(creds));
// } catch (error) {
//   console.log("No Creds found");
// }

app.get("/auth/google", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/userinfo.profile"],
  });
  res.redirect(url);
});

app.get("/google/redirect", async (req, res) => {
  const { code } = req.query;
  const { tokens } = await oauth2Client.getToken(code);
  oauth2Client.setCredentials(tokens);
  global.oauth2Client = oauth2Client;
  fs.writeFileSync("creds.json", JSON.stringify(tokens));
  res.send("Success");
});

// view engine setup
// app.set('views', path.join(__dirname, 'views'));
app.set("view engine", "jade");
app.use(cors());

app.use(logger("dev"));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, "views")));

app.use("/", mainRoute);

// catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(createError(404));
});

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render("error");
});

module.exports = app;
