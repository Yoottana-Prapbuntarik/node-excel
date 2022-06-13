const express = require('express')
const app = express()

var excel = require('excel4node');
// Create a new instance of a Workbook class
var workbook = new excel.Workbook();

var style = workbook.createStyle({
  font: {
    name: 'Angsana New',
    color: '#000000',
    size: 18,
  },
  alignment: {
    horizontal: ['center'],
  },
    fill: {
    type: 'pattern',
    patternType: 'solid',
    bgColor: '#fabf90',
    fgColor: '#fabf90',
  },
});

var styleII = workbook.createStyle({
  alignment: {
    horizontal: ['center'],
  },
  font: {
    name: 'Angsana New',
    color: '#000000',
    size: 18,
  },
})
// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet 1');
// heading
worksheet.cell(1, 1, 1, 43, true).string('แบบทะเบียนสมาชิกสามัญขององค์การคนพิการแต่ล่ะประเภท ระบุชื่อองค์กร').style(style)

worksheet.cell(2, 1, 2, 2, true).string('ลำดับที่').style(styleII)
worksheet.cell(2, 3, 2, 4, true).string('คำนำหน้า').style(styleII)
worksheet.cell(2, 5, 2, 6, true).string('ชื่อ').style(styleII)
worksheet.cell(2, 7, 2, 8, true).string('นามสกุล').style(styleII)
worksheet.cell(2, 9, 2, 10, true).string('วัน/เดือน/ปี').style(styleII)
worksheet.cell(2, 11, 2, 12, true).string('เกิด	อายุ (ปี)').style(styleII)
worksheet.cell(2, 13, 2, 15, true).string('หมายเลขบัตรประชาชน').style(styleII)
// heading
worksheet.cell(2, 16, 2, 29, true).string('ที่อยู่').style(styleII)

worksheet.cell(3, 16, 3, 17, true).string('บ้านเลขที่').style(styleII)
worksheet.cell(3, 18, 3, 19, true).string('หมู่ที่').style(styleII)
worksheet.cell(3, 20, 3, 21, true).string('อาคาร/หมู่บ้าน').style(styleII)
worksheet.cell(3, 22, 3, 23, true).string('ตรอก/ซอย').style(styleII)
worksheet.cell(3, 24, 3, 25, true).string('ถนน').style(styleII)
worksheet.cell(3, 26, 3, 27, true).string('ตำบล/แขวง').style(styleII)
worksheet.cell(3, 28, 3, 29, true).string('อำเภอ/เขต').style(styleII)
worksheet.cell(3, 30, 3, 31, true).string('จังหวัด').style(styleII)
worksheet.cell(2, 32, 2, 33, true).string('เบอร์โทรศัพท์').style(styleII)

// // heading
worksheet.cell(2, 34, 2, 41, true).string('สถานผู้สมัคร').style(styleII)

worksheet.cell(3, 34, 3, 35, true).string('สถานะ').style(styleII)
worksheet.cell(3, 36, 3, 37, true).string('ประเภทความพิการ').style(styleII)
worksheet.cell(3, 38, 3, 39, true).string('ชื่อ-สกุล คนพิการ').style(styleII)
worksheet.cell(3, 40, 3, 41, true).string('หมายเลขบัตรคนพิการ').style(styleII)
worksheet.cell(2, 42, 2, 43, true).string('วันเดือนปีที่สมัคร').style(styleII)

// mockup data
for(let i=0; i<50; i++) {
  // begin cell 4 + i <<< i refer to new row
  // content
  worksheet.cell(4+i, 3, 4+i, 4, true).string('นาย').style(styleII)
  worksheet.cell(4+i, 1, 4+i, 2, true).string(`${i+1}`).style(styleII)
  worksheet.cell(4+i, 5, 4+i, 6, true).string('ทดสอบ').style(styleII)
  worksheet.cell(4+i, 7, 4+i, 8, true).string('ระบบ').style(styleII)
  worksheet.cell(4+i, 9, 4+i, 10, true).string('30/11/2022').style(styleII)
  worksheet.cell(4+i, 11, 4+i, 12, true).string('15 (ปี)').style(styleII)
  worksheet.cell(4+i, 13, 4+i, 15, true).string('1101500910101').style(styleII)

  worksheet.cell(4+i, 16, 4+i, 17, true).string('99').style(styleII)
  worksheet.cell(4+i, 18, 4+i, 19, true).string('-').style(styleII)
  worksheet.cell(4+i, 20, 4+i, 21, true).string('อมรชัย9').style(styleII)
  worksheet.cell(4+i, 22, 4+i, 23, true).string('พระราม2 36').style(styleII)
  worksheet.cell(4+i, 24, 4+i, 25, true).string('พระราม2').style(styleII)
  worksheet.cell(4+i, 26, 4+i, 27, true).string('บางมด').style(styleII)
  worksheet.cell(4+i, 28, 4+i, 29, true).string('จอมทอง').style(styleII)
  worksheet.cell(4+i, 30, 4+i, 31, true).string('กรุงเทพมหานคร').style(styleII)
  worksheet.cell(4+i, 32, 4+i, 33, true).string('0961238923').style(styleII)

  worksheet.cell(4+i, 34, 4+i, 35, true).string('คนพิการ').style(styleII)
  worksheet.cell(4+i, 36, 4+i, 37, true).string('พิการทางสายตา, ออทิสติก').style(styleII)
  worksheet.cell(4+i, 38, 4+i, 39, true).string('ออทิสติก ทดสอบระบบ').style(styleII)
  worksheet.cell(4+i, 40, 4+i, 41, true).string('1234567891011').style(styleII)
  worksheet.cell(4+i, 42, 4+i, 43, true).string('30/10/2022').style(styleII)
}

workbook.write('Excel.xlsx');
app.get('/', function (req, res) {
  res.send('Hello World')
})

app.listen(3009, ()=> {
  console.log('running on the port ', '3009')
})
