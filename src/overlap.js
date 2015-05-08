var xlsx = require('node-xlsx');
var file_system = require('fs');

var sum = function(arr) {
	var _ret = 0;
	for (var i = 0; i < arr.length; i++) {
		var num = arr[i];
		if (typeof num == "number") _ret += num;
	}
	return _ret;
};

var calcOverlap = function(adv1, adv2) {
	var len = adv1.length,
		_denominator1 = 0,
		_denominator2 = 0,
		_molecule = 0;
	for (var i = 0; i < len; i++) {
		var v1 = adv1[i],
			v2 = adv2[i];
		_molecule += v1*v2;
		_denominator1 += v1*v1;
		_denominator2 += v2*v2;
	}
	return _molecule / Math.sqrt(_denominator1*_denominator2);
};

var calcLevinsWidth = function(adv) {
	var sigma = 0, sumadv = 0, r = adv.length;
	for (var i = 0; i < r; i++) {
		sumadv += adv[i];
	}
	for (var i = 0; i < r; i++) {
		var advp = adv[i]/sumadv,
			_ans = advp * Math.log(advp);
		sigma += isNaN(_ans) ? 0 : _ans;
	}
	return -sigma;
};

var calcHurlbertWidth = function(adv) {
	var sigma = 0, sigmap2 = 0, sumadv = 0, r = adv.length;
	for (var i = 0; i < r; i++) {
		sumadv += adv[i];
	}
	for (var i = 0; i < r; i++) {
		var advp = adv[i]/sumadv;
		sigmap2 += advp * advp;
	}
	return (1/sigmap2 - 1) / (r - 1);
};

var makeZeroArray = function(len) {
	var _ret = [];
	for (var i = 0; i < len; i++) _ret.push(0);
	return _ret;
};

var util = {
	xlsx: xlsx,
	readData: function(data, names, advantages) {
		var records = {},
			cols = names.length,
			c = 0,
			r = 0;
		for (r = 0; r < data.length; r++) {
			var row = data[r];
			for (c = 0; c < cols; c++) {
				var name = row[names[c]];
				if (typeof name != "string") continue;
				var advantage = parseFloat(row[advantages[c]]);
				if (isNaN(advantage)) continue;
				if (!records[name]) records[name] = makeZeroArray(cols);
				records[name][c] = advantage;
			}
		}
		return records;
	},

	readTable: function(filename) {
		var data = xlsx.parse(filename);
		for (var i = 0; i < data.length; i++) {
			var item = data[i];
			item.data = util.readData(item.data, [2,5,8,11], [3,6,9,12]);
		}
		return data;
	},

	rankAdvantages: function(records) {
		var arr = [];
		for (var k in records) {
			var advantages = records[k];
			arr.push({
				name: k,
				adv: sum(advantages),
				levinsWidth: calcLevinsWidth(advantages),
				hurlbertWidth: calcHurlbertWidth(advantages)
			});
		}
		arr = arr.sort(function(a,b) {
			return b.adv - a.adv;
		});
		return arr;
	},

	overlapMatric: function(records) {
		var matric = {};
		for (var k1 in records) {
			var advantages1 = records[k1];
			var mr = matric[k1] = {};
			for (var k2 in records) {
				var advantages2 = records[k2];
				if (matric[k2] && (matric[k2][k1] !== undefined)) mr[k2] = matric[k2][k1];
				else mr[k2] = calcOverlap(advantages1, advantages2);
			}
		}
		return matric;
	},

	formStatisticsTable: function(statistics, options) {
		// options:
		// -- header_row
		// -- sortby
		options = options || {};
		var tableData = [
			options.header_row || ['物种', 'Levins宽度', 'Hurlbert宽度']
		];
		var _statistics = statistics;
		if (options.sortby) {
			_statistics = statistics.sort(function(a,b) {
				return b[options.sortby] - a[options.sortby];
			});
		}
		for (var i = 0; i < _statistics.length; i++) {
			var item = _statistics[i];
			// if (isNaN(item.levinsWidth) ||  
			// 	isNaN(item.hurlbertWidth)) 
			// 	continue;
			tableData.push([
				item.name,
				item.levinsWidth,
				item.hurlbertWidth
			]);
		}
		return tableData;
	},

	formOverlapTable: function(matric) {
		var tableData = [],
			keyArr = [],
			header_row = [undefined];
		for (var k in matric) {
			keyArr.push(k);
			header_row.push(k);
		}
		tableData.push(header_row);
		for (var i = 0; i < keyArr.length; i++) {
			var k1 = keyArr[i];
				row = [k1];
			for (var j = 0; j < keyArr.length; j++) {
				var k2 = keyArr[j];
				if (j < i) row.push(undefined);
				else row.push(matric[k1][k2]);
			}
			tableData.push(row);
		}
		return tableData;
	},

	outputData: function(filename, data, callback) {
		var file_bin = xlsx.build(data);
		file_system.writeFile(filename, file_bin, callback);
	},
	outputDataSync: function(filename, data) {
		var file_bin = xlsx.build(data);
		file_system.writeFileSync(filename, file_bin);
	}
}

module.exports = util;