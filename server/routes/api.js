const express = require('express');
const router = express.Router();
const transformController = require('../controllers/transformController');

// 转换接口
router.post('/transform', transformController.handleTransform.bind(transformController));

module.exports = router; 