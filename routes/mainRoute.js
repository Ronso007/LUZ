var express = require("express");
var router = express.Router();

// Require route modules.
var schedule_route = require("./ScheduleRoute");

/* GET home page. */
router.get("/", function (req, res, next) {
  res.sendFile("index", { root: "./" });
});

router.use("/api/schedule", schedule_route);

module.exports = router;
