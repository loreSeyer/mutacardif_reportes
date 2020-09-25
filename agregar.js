var mysql = require('mysql');
const readXlsxFile = require('read-excel-file/node');
var pdf = require('html-pdf');
const Path          = require('path');
const fs            = require('graceful-fs');
const moment            = require('moment');
const async          = require('async');
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
  
    readline.question(`tipo: `, tip => {
      tipo=tip;
      console.log(tipo);
      // readline.close()
      return callback(null,"Final entrada");
    })
  
}

function entrada5(callback){
    readline.question(`parte: `, part => {
      parte=part;
      console.log(parte);
      readline.close()
      return callback(null,"Final entrada");
    })
  }



function create(callback){

    con.connect(function(err) {
        if (err) throw err;
        console.log("Connected!");
      //  
          if(tipo==="ACT"){
            if(producto==="INGRESOS" || producto==="VALORES"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255), POLIZA varchar(255), NOMBRETIT varchar(255) ,APPATERNO varchar(255),PLAN varchar(255), CODIGO varchar(255), CONTROL varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA, POLIZA , NOMBRETIT, APPATERNO, PLAN, CODIGO, CONTROL) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }

            if(producto==="PLENITUD"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255), POLIZA varchar(255), NOMBRETIT varchar(255) ,APPATERNO varchar(255),PLAN varchar(255), CODIGO varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA, POLIZA , NOMBRETIT, APPATERNO, PLAN, CODIGO) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
          }
          if(tipo==="POLIZA"){
            if(producto==="INGRESOS" || producto==="VALORES"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), FOLIO VARCHAR(255), ID varchar(255), AP_PATERNO varchar(255) ,NOMBRE varchar(255),FECHA_INICIO varchar(255), FECHA_FIN varchar(255), CODIGO varchar(255), ARCHIVO varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, FOLIO, ID , AP_PATERNO, NOMBRE, FECHA_INICIO, FECHA_FIN, CODIGO, ARCHIVO) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }

            if(producto==="PLENITUD"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255), POLIZA varchar(255), AP_PATERNO varchar(255) ,NOMBRE varchar(255),FECHA_INICIO varchar(255), FECHA_FIN varchar(255), FOLIO varchar(255), ARCHIVO varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA, POLIZA , AP_PATERNO, NOMBRE, FECHA_INICIO, FECHA_FIN, FOLIO, ARCHIVO) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
            if(producto==="COPPEL"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), FOLIO VARCHAR(255), NOMBRE varchar(255), AP_PATERNO varchar(255) ,NO_CERTIFICADO varchar(255),FECHA_INICIO varchar(255), FECHA_FIN varchar(255), ARCHIVO varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, FOLIO, NOMBRE, AP_PATERNO, NO_CERTIFICADO, FECHA_INICIO, FECHA_FIN, ARCHIVO) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
            if(producto==="BANORTE_TMK"){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255),NO_CERTIFICADO varchar(255), NO_POLIZA varchar(255), NOMBRE varchar(255),FECHA_INICIO varchar(255), FECHA_FIN varchar(255))";
              

              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA,NO_CERTIFICADO,NO_POLIZA, NOMBRE,FECHA_INICIO,FECHA_FIN) VALUES ?";
                        
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }

            if(producto==="BANORTE_TMK_INB"){
              // var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255),NO_CERTIFICADO varchar(255), NO_POLIZA varchar(255), NOMBRE varchar(255))";
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), REMESA VARCHAR(255),NO_CERTIFICADO varchar(255), NO_POLIZA varchar(255), NOMBRE varchar(255), FECHA_INICIO varchar(255), FECHA_FIN varchar(255))";

              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        // var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA,NO_CERTIFICADO,NO_POLIZA, NOMBRE) VALUES ?";
                        var sql1 = "INSERT INTO "+nombre+" (CORREO, REMESA,NO_CERTIFICADO,NO_POLIZA, NOMBRE, FECHA_INICIO, FECHA_FIN) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
          }

          if(tipo==="API"){
            if(producto==="COPPEL"){
              var sql = "CREATE TABLE "+nombre+" (ID VARCHAR(255), CORREO VARCHAR(255), FECHA varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (ID, CORREO, FECHA) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
          }

          if(tipo==="ESPECIAL"){

            if(producto==="INGRESOS" || producto==="MOMENTOS" || producto==="VALORA" || producto==="VALORES" ){
              var sql = "CREATE TABLE "+nombre+" (CORREO VARCHAR(255), NOMBRE_ASEG VARCHAR(255), INICIO_VIGENCIA varchar(255), POLIZA_2 varchar(255))";
              con.query(sql, function (err, result) {
                  if (err) throw err;
                  console.log("Table "+nombre+" created");
    
                    readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+tipo+'/'+nombre+'.xlsx').then((rows) => {
                        // console.log(rows);
    
                        var sql1 = "INSERT INTO "+nombre+" (CORREO,NOMBRE_ASEG,INICIO_VIGENCIA, POLIZA_2) VALUES ?";
                        con.query(sql1, [rows], function(err) {
                            if (err) throw err;
                            console.log("Information inserted!");
                            con.end();
                        });
                    })
    
              });
            }
          }
      });

}