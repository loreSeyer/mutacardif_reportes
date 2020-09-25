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
        create(callback);
    }
]);

let nombre;
let producto;
let mes;
let parte;

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
      readline.close()
      return callback(null,"Final entrada");
    })
  }




let pdfs=[];

function create(callback){

    con.connect(function(err) {
        if (err) throw err;
        console.log("Connected!");
      //  
          var sql = "CREATE TABLE "+nombre+" (campa VARCHAR(255), email VARCHAR(255), fecha_envio varchar(255), estatus varchar(255) ,fecha_estatus varchar(255), detalle varchar(500))";
          con.query(sql, function (err, result) {
              if (err) throw err;
              console.log("Table "+nombre+" created");

                readXlsxFile('./bases/'+mes+'2020/'+mes+'-'+parte+'/'+producto+'/'+nombre+'.xlsx').then((rows) => {
                    // console.log(rows);

                    var sql1 = "INSERT INTO "+nombre+" (campa, email, fecha_envio , estatus, fecha_estatus, detalle) VALUES ?";
                    con.query(sql1, [rows], function(err) {
                        if (err) throw err;
                        console.log("Information inserted!");
                        con.end();
                    });
                })

          });
      });

}