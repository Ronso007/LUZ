// Require route modules.
var schedule_action = require("../actions/ExcelActions/scheduleAction");

exports.schedule_week_with_staff = async function (req, res) {
  try {
    const weekNumber = req.query.Week;
    const Cycle = req.query.Cycle;
    res.send(JSON.stringify(await schedule_action.readWeek(weekNumber, Cycle)));
  } catch (error) {}
};

exports.schedule_get_num_of_weeks = async function (req, res) {
  const Cycle = req.query.Cycle;

  let numOfWeeks = await schedule_action.getNumOfWeeks(Cycle);

  res.send(JSON.stringify(numOfWeeks));
};
