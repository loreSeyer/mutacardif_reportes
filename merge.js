var mysql = require('mysql');
const readXlsxFile = require('read-excel-file/node');
var pdf = require('html-pdf');
const Path          = require('path');
const fs            = require('graceful-fs');
const moment            = require('moment');
const async          = require('async');
const xlsx = require("xlsx");
const readline = require('readline').createInterface({
  input: process.stdin,
  output: process.stdout
})

var con = mysql.createConnection({
  host: "localhost",
  user: "root",
  password: "SEyer1428**",
  database: "cardif"
});

// let name="coppel1_agosto2020";

async.series([
    function(callback) {
        entrada4(callback);
    },
    function(callback) {
        entrada5(callback);
    },
    function(callback) {
        entrada7(callback);
    },
    function(callback) {
        create(callback);
    }
]);

let nombre1;
let nombre2;
let nombre3;
let tipo;
let producto;
let mes;
let ciclo;


function entrada4(callback){
  
    readline.question(`mes: `, mes1 => {
      mes=mes1;
      console.log(mes);
      // readline.close()
      return callback(null,"Final entrada");
    });

}

function entrada5(callback){
  
    readline.question(`producto: `, prod => {
      producto=prod;
      console.log(producto);
      // readline.close()
      return callback(null,"Final entrada");
    });

}


function entrada7(callback){
    readline.question(`tipo: `, tip => {
      tipo=tip;
      console.log(tipo);
      readline.close()
      return callback(null,"Final entrada");
    })
  }



function create(callback){
    let cont_ok=0;
    let cont_error=0;
    let data=[];
    let data1=[];
    let direc=Path.join(__dirname+'/listas/GENERAL/'+tipo+'/'+mes+'');
    try {
        fs.statSync(direc);
        console.log('file or directory exists');
        ciclo=2;
    }
    catch (err) {
      if (err.code === 'ENOENT') {
        
        fs.mkdirSync(direc);
        ciclo=1;
      }
    }

    
    let url;
    // console.log("se abre archivo y luego se agregan datos");
    if(tipo==="ACT"){
        url='./listas/'+producto+'/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx';
    }
    if(tipo==="POLIZA"){
        url='./listas/'+producto+'/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-CARATULA-'+mes+'.xlsx'
    }

    var newWB = xlsx.utils.book_new()
    if(ciclo===1){
        

        readXlsxFile(url).then((rows) => {
    
            // console.log(rows);
            let encabezados=rows[0];
            for (const row of rows) {
                let i=0;
    
                let fila={};
                let dataa=row;
                for (const dat of dataa) {
                    let name=encabezados[i];
                    
                    if(dat!==null){
                        fila[name]=""+dat;   
                    }
                    else{
                        fila[name]="";
                    }
                    i++;
                }
                data.push(fila);
            }
    
            data.splice(0,1);
            
            var newWS = xlsx.utils.json_to_sheet(data);
            xlsx.utils.book_append_sheet(newWB,newWS,"MAIL_"+tipo)//workbook name as param
            xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
    
        })
    }
    else{
        console.log("se abre archivo y luego se agregan datos 2");
    

        readXlsxFile('./listas/GENERAL/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx').then((rows) => {
    
            let encabezados=rows[0];
            for (const row of rows) {
                let i=0;
    
                let fila={};
                let dataa=row;
                for (const dat of dataa) {
                    let name=encabezados[i];
                    
                    if(dat!==null){
                        fila[name]=""+dat;   
                    }
                    else{
                        fila[name]="";
                    }
                    i++;
                }
                data.push(fila);
                
            }
    
            data.splice(0,1);
            
            readXlsxFile(url).then((rows) => {
                // console.log(data);
        
                // console.log(rows);
                let encabezados=rows[0];
                for (const row of rows) {
                    let i=0;
        
                    let fila={};
                    let dataa=row;
                    for (const dat of dataa) {
                        let name=encabezados[i];
                        
                        if(dat!==null){
                            fila[name]=""+dat;   
                        }
                        else{
                            fila[name]="";
                        }
                        i++;
                    }
                        data1.push(fila);
                    
                }
        
                data1.splice(0,1);

                data=data.concat(data1);
                
                var newWS = xlsx.utils.json_to_sheet(data);
                xlsx.utils.book_append_sheet(newWB,newWS,"MAIL_"+tipo)//workbook name as param
                xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
        
            })
    
        })
    }
}