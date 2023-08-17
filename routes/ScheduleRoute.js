var express = require("express");
var router = express.Router();

// Require controller modules.
var schedule_controller = require("../controllers/ScheduleController");

// Schedule Routes

// GET
router.get("/", schedule_controller.schedule_week_with_staff);

router.get("/weeks", schedule_controller.schedule_get_num_of_weeks);

module.exports = router;
