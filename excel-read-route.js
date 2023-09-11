const { readExcelController } = require('./excelController');

const express = require('express');

const router = express.Router();

router.get('/calc', readExcelController);

module.exports = router;
