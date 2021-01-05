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

// let name="coppel1_agosto2020";

async.series([
    function(callback) {
        entrada(callback);
    },
    function(callback) {
        entrada2(callback);
    },
    function(callback) {
        entrada3(callback);
    },
    function(callback) {
        entrada4(callback);
    },

    function(callback) {
        entrada5(callback);
    },
    function(callback) {
        create(callback);
    }
]);

let nombre;
let producto;
let mes;
let parte;
let tipo;
let año="2021";

function entrada(callback){
  readline.question(`Nombre de base: `, name => {
    nombre=name;
    console.log(nombre);
    // readline.close()
    return callback(null,"Final entrada");
  });
}

function entrada2(callback){
  
    readline.question(`Mes de reporte: `, mes1 => {
      mes=mes1;
      console.log(mes);
      // readline.close()
      return callback(null,"Final entrada");
    });

}

function entrada3(callback){
  
    readline.question(`producto: `, prod => {
      producto=prod;
      console.log(producto);
      // readline.close()
      return callback(null,"Final entrada");
    })
  
}

function entrada4(callback){
    readline.question(`parte: `, part => {
      parte=part;
      console.log(parte);
      return callback(null,"Final entrada");
    })
  }

  function entrada5(callback){
    readline.question(`tipo: `, tip => {
      tipo=tip;
      console.log(tipo);
      readline.close()
      return callback(null,"Final entrada");
    })
  }




let pdfs=[];

function create(callback){

    let content=[];
    fs.readFile('./bases/A'+año+'/'+mes+'/PARTE-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.json', 'utf-8', function (err, fileContents) { 
        if (err) throw err; 
        // console.log(JSON.parse(fileContents)); 

        let direc=Path.join(__dirname+'/listas/ENTREGABLES/REPORTES-'+año+'/'+producto+'/'+tipo+'/'+mes+'');
            try {
                fs.statSync(direc);
                console.log('file or directory exists');
            }
            catch (err) {
            if (err.code === 'ENOENT') {
                
                fs.mkdirSync(direc);
            }
        }

        if(producto==="GENERAL"){
            let data= JSON.parse(fileContents);
            for (const item of data) {
                let dat={};
                dat.TELEFONO=item.number;
                dat.FECHA=item.fecha;
                dat.ESTADO=item.status;
                dat.BASE_DE_DATOS=item.database;
                dat.REMESA=item.data_info.REMESA;
                dat.POLIZA=item.data_info.POLIZA;

                if(item.status==="Error"){
                    dat.VALIDO="";
                    dat.INVALIDO="1"

                }
                else{
                    if(item.status==="Exitoso"){
                        dat.VALIDO="1"
                        dat.INVALIDO="";
                    }
                }

                content.push(dat);  
            }
            console.log(content);
        }

        if(producto==="COPPEL"){
            let data= JSON.parse(fileContents);
            for (const item of data) {
                let dat={};
                dat.ID=item._id;
                dat.TELEFONO=item.number;
                dat.FECHA=item.fecha;
                dat.ESTADO=item.status;
                content.push(dat);  
            }
            console.log(content);
        }

        var newWB = xlsx.utils.book_new();
        var newWS = xlsx.utils.json_to_sheet(content);
        xlsx.utils.book_append_sheet(newWB,newWS,"SMS")//workbook name as param
        xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-SMS-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
    }); 

    

}