var xlsx = require('node-xlsx');
var util = require('./overlap');

var data = util.readTable('/Users/zhangcosmo/Downloads/test.xlsx');

// console.log(util.overlapMatric(data[0].data));
util.outputDataSync('/Users/zhangcosmo/Downloads/test_out.xlsx', [
	{
		name: '1',
		data: util.formStatisticsTable(util.rankAdvantages(data[0].data), {
			sortby: 'levinsWidth'
		})
	}
]);