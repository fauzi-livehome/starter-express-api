const xlsx = require('xlsx');
const xlsx_calc = require('xlsx-calc');

module.exports = {
  readExcelController: (req, res) => {
    const workbook = xlsx.readFile('./Quiz_Question.xlsx', { cellFormula: true });
    let workbook_sheet = workbook.SheetNames;
    let workbook_update_amount = xlsx.utils.sheet_add_aoa(workbook.Sheets[workbook_sheet[0]], [[req.query.amt]], { origin: "B4" });
    let workbook_update_duration = xlsx.utils.sheet_add_aoa(workbook.Sheets[workbook_sheet[0]], [[req.query.dur]], { origin: "B13" });
    xlsx_calc(workbook);
    xlsx_calc(workbook, { continue_after_error: true, log_error: true })
    let workbook_response = xlsx.utils.sheet_to_json(
      workbook.Sheets[workbook_sheet[0]]
    );
    res.status(200).send({
      message: workbook_response,
    });
  },
};
