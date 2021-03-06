const express=require("express");
const app=express()
const bodyParser=require("body-parser")
const excel = require('excel4node');
const fs = require('fs');
const path = require('path');
let dirPath = __dirname;
dirPath = dirPath.replace(/\\/gm, "/") + "/";



const port=9000;
let workbook = new excel.Workbook();



app.use(bodyParser.urlencoded());
app.get("/excelFile",(req,res)=>{
  try{
let worksheet = workbook.addWorksheet('Sheet 1');
let style = workbook.createStyle({
  font: {
    color: '#FF0800',
    size: 12
  },
  numberFormat: '#,##0.00; (#,##0.00); -'
});

// Set value of cell A1 to 10 as a number 
worksheet.cell(1,1).number(10).style(style);

// Set value of cell B1 to 20 as a number 
worksheet.cell(1,2).number(20).style(style);

// Set value of cell C1 to a formula 
worksheet.cell(1,3).formula('A1 + B1').style(style);


workbook.write('Excel.xlsx');
return res.download(dirPath+'/Excel.xlsx');
  }
  catch(ex){
   
    res.status(500).json("internal server error",ex)
  }
});



app.listen(port,()=>{
    console.log(`port listen on http://localhost:${port}`);
});