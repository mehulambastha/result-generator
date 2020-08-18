/*
	Node.Js program to generate Excel File from Text File sent by CBSE
	Made by: Mehul Ambastha
	contact: mehul.213amb@gmail.com

	NOTE: You must have node installed, be in the directory of the the file, and then run 		the file by running 'node result10.js' in the terminal
	NOTE: The SAMPLE_FILE.txt must be in the same folder, and result-10.xlsx will be 		  generated in the same folder
*/

const fs = require('fs');
const excel = require('excel4node');

let workbook = new excel.Workbook({
	defaultFont: {
		name: 'Arial',
	},
});

let worksheet = workbook.addWorksheet('Result-sheet');

let headerStyle = workbook.createStyle({
	font: {
		color: '#000000',
		size: 12,
	},
	numberFormat: '$#, ##0.00; ($#,##0.00); -',
});
let dataStyle = workbook.createStyle({
	font: {
		color: '#060606',
		size: 10.5,
	},
	alignment: {
		vertical: 'center',
	},
});

let headers = [
	'Roll no',
	'Gender',
	'Student Name',
	'English',
	'Hindi',
	'Maths-Basic',
	'Science',
	'Social Science',
	'Computer',
	'Maths-Standard',
	'Painting',
	'Music',
	'Sanskrit',
	'Result',
];

for (let i = 1; i <= headers.length; i++) {
	worksheet
		.cell(1, i)
		.string(headers[i - 1])
		.style(headerStyle);
}

//Reading the sample-result file
let readTxtFile = fs.createReadStream('./SAMPLE_FILE.txt');
let lines;
let data = '';

readTxtFile
	.on('data', text => {
		data += text;
		lines = data.split(/\r?\n/);
	})
	.on('end', () => {
		lines.pop();
		processData(lines);
	});

function processData(lines) {
	let row = 2;
	let count = 1;
	for (let k = 0; k < lines.length; k++) {
		line = lines[k].split(/\s+/);

		if (count % 2 === 1) {
			line.pop();
			var name;
			var roll = line[0];
			var gender = line[1];
			var subjectWiseData = new Map();
			var subjectsArray = [];
			var marksPerChild = [];
			let iterable;
			var result = line[line.length - 1];
			if (isNaN(line[4])) {
				name = `${line[2]} ${line[4]}`;
				iterable = 5;
			} else {
				name = `${line[2]} ${line[3]}`;
				iterable = 4;
			}

			for (let i = iterable; i < line.length - 1; i++) {
				subjectsArray.push(line[i]);
				subjectWiseData.set(line[i], '');
			}
		} else {
			line.shift();
			for (let i = 0; i < line.length; i++) {
				if (i % 2 == 0) {
					marksPerChild.push(line[i]);
				}
			}
			for (let j = 0; j < marksPerChild.length; j++) {
				subjectWiseData.set(subjectsArray[j], marksPerChild[j]);
			}
			writeToExcel(roll, name, gender, subjectWiseData, result, row);
			row += 1;
		}
		count += 1;
	}
}

var fileName = 'result-10';
var filePath = `./${fileName}.xlsx`;

function writeToExcel(roll, name, gender, subjectData, result, row) {
	worksheet.cell(row, 1).string(roll).style(dataStyle);
	worksheet.cell(row, 2).string(gender).style(dataStyle);
	worksheet.cell(row, 3).string(name).style(dataStyle);
	worksheet.cell(row, 14).string(result).style(dataStyle);
	for (let code of subjectData.keys()) {
		if (code == '184') {
			worksheet.cell(row, 4).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '085') {
			worksheet.cell(row, 5).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '241') {
			worksheet.cell(row, 6).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '086') {
			worksheet.cell(row, 7).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '087') {
			worksheet.cell(row, 8).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '165') {
			worksheet.cell(row, 9).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '041') {
			worksheet.cell(row, 10).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '049') {
			worksheet.cell(row, 11).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '034') {
			worksheet.cell(row, 12).string(subjectData.get(code)).style(dataStyle);
		} else if (code == '122') {
			worksheet.cell(row, 13).string(subjectData.get(code)).style(dataStyle);
		}
	}
	workbook.write(filePath);
}

if (!fs.existsSync(filePath)) {
	console.log(
		`The file has been generated with the name of ${fileName}.xlsx . Please check the folder...`
	);
}
