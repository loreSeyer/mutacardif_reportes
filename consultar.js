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
        entrada6(callback);
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

function entrada(callback){
  readline.question(`Nombre de tabla 1: `, name => {
    nombre1=name;
    console.log(nombre1);
    // readline.close()
    return callback(null,"Final entrada");
  });
}

function entrada2(callback){
  
    readline.question(`Nombre de tabla 2: `, name2 => {
      nombre2=name2;
      console.log(nombre2);
      // readline.close()
      return callback(null,"Final entrada");
    });

}

function entrada3(callback){
  
    readline.question(`Nombre de tabla 3: `, name3 => {
      nombre3=name3;
      console.log(nombre3);
      // readline.close()
      return callback(null,"Final entrada");
    });

}

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

function entrada6(callback){
  
    readline.question(`Fecha: `, fech => {
      fecha=fech;
      console.log(fecha);
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
    let direc=Path.join(__dirname+'/listas/'+producto+'/'+tipo+'/'+mes+'');
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

    con.connect(function(err) {
        if (err) throw err;
        console.log("Connected!");
        let open;
        let delivered;
        let error;
        let click;

        if(tipo==="ACT"){
              let data=[];

            if(producto==="INGRESOS" || producto==="VALORES"){
                var sql = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.poliza,mo.fecha_estatus,mo.estatus,mog.NOMBRETIT,mog.APPATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza";

                var sql1="SELECT mo2.REMESA,mo1.email,mo2.poliza,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRETIT,mo2.APPATERNO from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza) group by mo2.poliza";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.poliza,er.fecha_estatus,er.estatus,mo.NOMBRETIT,mo.APPATERNO FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.poliza;";

                var sql3 = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.poliza,mo.fecha_estatus,mo.estatus, mog.NOMBRETIT,mog.APPATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.poliza";
            }
            if(producto==="PLENITUD"){

                var sql = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.poliza,mo.fecha_estatus,mo.estatus, mog.NOMBRETIT,mog.APPATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza";

                var sql1="SELECT mo2.REMESA,mo1.email,mo2.poliza,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRETIT,mo2.APPATERNO from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza) group by mo2.poliza;";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.poliza,er.fecha_estatus,er.estatus,mo.NOMBRETIT,mo.APPATERNO FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.poliza";

                var sql3 = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.poliza,mo.fecha_estatus,mo.estatus, mog.NOMBRETIT,mog.APPATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.poliza";
            }
            
                con.query(sql, function (err, results) {
                    if(results!==undefined){
                        cont_ok=results.length;
                    }
                    open=results;
                    console.log("llega a open");

                    if (err){
                        console.log(err);
                    };
                    con.query(sql1, function(err,results1){
                        console.log("llega a delivery");
                        if (err){
                            console.log(err);
                        };
                        
                        if(results1!==undefined){
                         
                            cont_ok=cont_ok+results1.length;
                            
                        }
                        delivered=results1;

                        con.query(sql2, function(err,results2){
                            if(results2!==undefined){
                                cont_error=results2.length;
                                
                            }
                            error=results2;

                            con.query(sql3,function(err,results3){
                                click=results3
                                console.log("crear excel");
                                var newWB = xlsx.utils.book_new()

                                if(ciclo===1){
                                    console.log("se crea el archivo de cero");
                                    
                                    if(open!==undefined){
                                        open.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            dat.REMESA=result.REMESA;
                                            dat.EMAIL=result.email;
                                            dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                            dat.POLIZA=result.poliza;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(delivered!==undefined){
                                        delivered.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            dat.REMESA=result.REMESA;
                                            dat.EMAIL=result.email;
                                            dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                            dat.POLIZA=result.poliza;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(error!==undefined){
                                        error.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            dat.REMESA=result.REMESA;
                                            dat.EMAIL=result.email;
                                            dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                            dat.POLIZA=result.poliza;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="";
                                            dat.INVALIDO="1";
                                            data.push(dat);
                                        });
                                    }

                                    if(click!==undefined){
                                        click.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            dat.REMESA=result.REMESA;
                                            dat.EMAIL=result.email;
                                            dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                            dat.POLIZA=result.poliza;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    // console.log(data);
                                    var newWS = xlsx.utils.json_to_sheet(data);
                                    xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                    xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
                                }
                                if(ciclo===2){
                                    console.log("se abre archivo y luego se agregan datos");
                                    let data_part=[];

                                    readXlsxFile('./listas/'+producto+'/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx').then((rows) => {
                    
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
                                        if(open!==undefined){
                                            open.forEach(result => {
                                                let dat={};
                                                // console.log(result);
                                                dat.FECHA=""+fecha;
                                                dat.PRODUCTO=""+producto;
                                                dat.REMESA=result.REMESA;
                                                dat.EMAIL=result.email;
                                                dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                                dat.POLIZA=result.poliza;
                                                dat.FECHA_ESTATUS=result.fecha_estatus;
                                                dat.ESTATUS=result.estatus;
                                                dat.REPETICIONES=result.repeticiones;
                                                dat.VALIDO="1";
                                                dat.INVALIDO="";
                                                data.push(dat);
                                            });
                                        }
                                        
                                        if(delivered!==undefined){
                                            delivered.forEach(result => {
                                                let dat={};
                                                // console.log(result);
                                                dat.FECHA=""+fecha;
                                                dat.PRODUCTO=""+producto;
                                                dat.REMESA=result.REMESA;
                                                dat.EMAIL=result.email;
                                                dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                                dat.POLIZA=result.poliza;
                                                dat.FECHA_ESTATUS=result.fecha_estatus;
                                                dat.ESTATUS=result.estatus;
                                                dat.REPETICIONES="";
                                                dat.VALIDO="1";
                                                dat.INVALIDO="";
                                                data.push(dat);
                                            });
                                        }
                                        
                                        if(error!==undefined){
                                            error.forEach(result => {
                                                let dat={};
                                                // console.log(result);
                                                dat.FECHA=""+fecha;
                                                dat.PRODUCTO=""+producto;
                                                dat.REMESA=result.REMESA;
                                                dat.EMAIL=result.email;
                                                dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                                dat.POLIZA=result.poliza;
                                                dat.FECHA_ESTATUS=result.fecha_estatus;
                                                dat.ESTATUS=result.estatus;
                                                dat.REPETICIONES="";
                                                dat.VALIDO="";
                                                dat.INVALIDO="1";
                                                data.push(dat);
                                            });
                                        }

                                        if(click!==undefined){
                                            click.forEach(result => {
                                                let dat={};
                                                // console.log(result);
                                                dat.FECHA=""+fecha;
                                                dat.PRODUCTO=""+producto;
                                                dat.REMESA=result.REMESA;
                                                dat.EMAIL=result.email;
                                                dat.NOMBRE=result.NOMBRETIT+" "+result.APPATERNO;
                                                dat.POLIZA=result.poliza;
                                                dat.FECHA_ESTATUS=result.fecha_estatus;
                                                dat.ESTATUS=result.estatus;
                                                dat.REPETICIONES=result.repeticiones;
                                                dat.VALIDO="1";
                                                dat.INVALIDO="";
                                                data.push(dat);
                                            });
                                        }
                                        // console.log(data);
                                        var newWS = xlsx.utils.json_to_sheet(data);
                                        xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                        xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});

                                    })
                                }
                                console.log("EXITOSOS: "+cont_ok);
                                console.log("ERROR: "+cont_error);
                                con.end();
                            })
                            
                        });
                    });
              });
            //}
          }

          if(tipo==="POLIZA"){
              console.log("se va a generar reporte de caratula");
            let data=[];

            if(producto==="INGRESOS" || producto==="VALORES" || producto==="INGRESOS-CAR"){
                var sql = "SELECT count(mo.estatus) repeticiones, mog.FOLIO,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.ID,mo.fecha_estatus,mo.estatus,mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.ID";

                var sql1="SELECT mo2.FOLIO,mo1.email,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo2.ID,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE, mo2.AP_PATERNO from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.ID) group by mo2.ID";

                var sql2="SELECT distinct mo.FOLIO,(er.email),mo.FECHA_INICIO,mo.FECHA_FIN,mo.ID,er.fecha_estatus,er.estatus, mo.NOMBRE, mo.AP_PATERNO FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.ID";

                var sql3 = "SELECT count(mo.estatus) repeticiones, mog.FOLIO,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.ID,mo.fecha_estatus,mo.estatus, mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.ID";
                
            }

            if(producto==="PLENITUD"){
                var sql = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.POLIZA,mo.fecha_estatus,mo.estatus, mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.POLIZA";

                var sql1="SELECT mo2.REMESA,mo1.email,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo2.POLIZA,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE, mo2.AP_PATERNO from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.POLIZA) group by mo2.POLIZA";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.FECHA_INICIO,mo.FECHA_FIN,mo.POLIZA,er.fecha_estatus,er.estatus, mo.NOMBRE, mo.AP_PATERNO FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.POLIZA";

                var sql3 = "SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.POLIZA,mo.fecha_estatus,mo.estatus, mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.POLIZA";

            }

            if(producto==="COPPEL"){
                var sql="SELECT count(mo.estatus) repeticiones,mog.FOLIO,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.NO_CERTIFICADO,mo.fecha_estatus,mo.estatus,mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO";

                var sql1="SELECT mo2.FOLIO,mo1.email,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo2.NO_CERTIFICADO,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE, mo2.AP_PATERNO from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO) and mo1.estatus='delivered' group by mo2.NO_CERTIFICADO";

                var sql2="SELECT distinct mo.FOLIO,(er.email),mo.FECHA_INICIO,mo.FECHA_FIN,mo.NO_CERTIFICADO,er.fecha_estatus,er.estatus, mo.NOMBRE, mo.AP_PATERNO FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.NO_CERTIFICADO;";

                var sql3="SELECT count(mo.estatus) repeticiones,mog.FOLIO,mo.email,mog.FECHA_INICIO,mog.FECHA_FIN,mog.NO_CERTIFICADO,mo.fecha_estatus,mo.estatus, mog.NOMBRE, mog.AP_PATERNO FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.NO_CERTIFICADO"

                
            }

            if(producto==="BANORTE_TMK"){
                var sql="SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.NO_CERTIFICADO,mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO";

                var sql1="SELECT mo2.REMESA,mo1.email, mo2.NO_CERTIFICADO,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO) and mo1.estatus='delivered' group by mo2.NO_CERTIFICADO";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.NO_CERTIFICADO,mo.FECHA_INICIO,mo.FECHA_FIN,er.fecha_estatus,er.estatus, mo.NOMBRE FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.NO_CERTIFICADO";

                var sql3="SELECT count(mo.estatus) repeticiones, mog.REMESA,mo.email,mog.NO_CERTIFICADO,mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.NO_CERTIFICADO"

                
            }

            if(producto==="BANORTE_TMK_INB"){
               
                //USAR APARTIR DE LA REMESA DE TMK-INBOUND
                var sql="SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.NO_CERTIFICADO,mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO";

                var sql1="SELECT mo2.REMESA,mo1.email, mo2.NO_CERTIFICADO,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO) and mo1.estatus='delivered' group by mo2.NO_CERTIFICADO";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.NO_CERTIFICADO,mo.FECHA_INICIO,mo.FECHA_FIN,er.fecha_estatus,er.estatus, mo.NOMBRE FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by er.email";

                var sql3="SELECT count(mo.estatus) repeticiones, mog.REMESA,mo.email,mog.NO_CERTIFICADO, mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.NO_CERTIFICADO"

                
            }
              con.query(sql, function (err, results) {
                  if(results!==undefined){
                      cont_ok=results.length;
                      
                  }
                  open=results;

                  if (err) throw err;

                  con.query(sql1, function(err,results1){
                      if(results1!==undefined){
                          cont_ok=cont_ok+results1.length;
                          
                      }
                      delivered=results1;

                      con.query(sql2, function(err,results2){
                          if(results2!==undefined){
                              cont_error=results2.length;
                              
                          }
                          error=results2;

                          con.query(sql3,function(err,results3){
                                 click=results3;

                                console.log("crear excel");
                                var newWB = xlsx.utils.book_new()

                                if(ciclo===1){
                                    console.log("se crea el archivo de cero");
                                    
                                    if(open!==undefined){
                                        open.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                            dat.EMAIL=result.email;
                                            result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                            // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                            result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                            result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                            producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(delivered!==undefined){
                                        delivered.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                            dat.EMAIL=result.email;
                                            result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                            // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                            result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                            result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                            producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(error!==undefined){
                                        error.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                            dat.EMAIL=result.email;
                                            result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                            // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                            result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                            result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                            producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="";
                                            dat.INVALIDO="1";
                                            data.push(dat);
                                        });
                                    }

                                    if(click!==undefined){
                                        click.forEach(result => {
                                            let dat={};
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+producto;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                            dat.EMAIL=result.email;
                                            result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                            // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                            result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                            result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                            producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                            dat.FECHA_ESTATUS=result.fecha_estatus;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                    
                                    // console.log(data);
                                    var newWS = xlsx.utils.json_to_sheet(data);
                                    xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                    xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-CARATULA-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
                                }
                                if(ciclo===2){
                                    console.log("se abre archivo y luego se agregan datos");
                                    let data_part=[];

                                    readXlsxFile('./listas/'+producto+'/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-CARATULA-'+mes+'.xlsx').then((rows) => {
                    
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
                                        if(open!==undefined){
                                                open.forEach(result => {
                                                    let dat={};
                                                    // console.log(result);
                                                    dat.FECHA=""+fecha;
                                                    dat.PRODUCTO=""+producto;
                                                    producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                                    dat.EMAIL=result.email;
                                                    result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                                    // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                                    result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                                    result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                                    producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                                    dat.FECHA_ESTATUS=result.fecha_estatus;
                                                    dat.ESTATUS=result.estatus;
                                                    dat.REPETICIONES=result.repeticiones;
                                                    dat.VALIDO="1";
                                                    dat.INVALIDO="";
                                                    data.push(dat);
                                                });
                                        }
                                        
                                            if(delivered!==undefined){
                                                delivered.forEach(result => {
                                                    let dat={};
                                                    // console.log(result);
                                                    dat.FECHA=""+fecha;
                                                    dat.PRODUCTO=""+producto;
                                                    producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                                    dat.EMAIL=result.email;
                                                    result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                                    // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                                    result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                                    result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                                    producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                                    dat.FECHA_ESTATUS=result.fecha_estatus;
                                                    dat.ESTATUS=result.estatus;
                                                    dat.REPETICIONES="";
                                                    dat.VALIDO="1";
                                                    dat.INVALIDO="";
                                                    data.push(dat);
                                                });
                                            }
                                        
                                            if(error!==undefined){
                                                error.forEach(result => {
                                                    let dat={};
                                                    // console.log(result);
                                                    dat.FECHA=""+fecha;
                                                    dat.PRODUCTO=""+producto;
                                                    producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                                    dat.EMAIL=result.email;
                                                    result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                                    // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                                    result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                                    result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                                    producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                                    dat.FECHA_ESTATUS=result.fecha_estatus;
                                                    dat.ESTATUS=result.estatus;
                                                    dat.REPETICIONES="";
                                                    dat.VALIDO="";
                                                    dat.INVALIDO="1";
                                                    data.push(dat);
                                                });
                                            }

                                            if(click!==undefined){
                                                click.forEach(result => {
                                                    let dat={};
                                                    dat.FECHA=""+fecha;
                                                    dat.PRODUCTO=""+producto;
                                                    producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:dat.REMESA=result.FOLIO;
                                                    dat.EMAIL=result.email;
                                                    result.AP_PATERNO!==undefined?dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO: dat.NOMBRE=result.NOMBRE;
                                                    // dat.NOMBRE=result.NOMBRE+" "+result.AP_PATERNO!==undefined?result.AP_PATERNO:"";
                                                    result.FECHA_INICIO!==undefined?dat.FECHA_INICIO=result.FECHA_INICIO: dat.FECHA_INICIO="";
                                                    result.FECHA_FIN!==undefined?dat.FECHA_FIN=result.FECHA_FIN:dat.FECHA_FIN="";
                                                    producto==="COPPEL" || producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="PLENITUD"?dat.POLIZA=result.POLIZA: dat.POLIZA=result.ID;
                                                    dat.FECHA_ESTATUS=result.fecha_estatus;
                                                    dat.ESTATUS=result.estatus;
                                                    dat.REPETICIONES=result.repeticiones;
                                                    dat.VALIDO="1";
                                                    dat.INVALIDO="";
                                                    data.push(dat);
                                                });
                                            }
                                        // console.log(data);
                                        var newWS = xlsx.utils.json_to_sheet(data);
                                        xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                        xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-CARATULA-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});

                                    })
                                }
                                console.log("EXITOSOS: "+cont_ok);
                                console.log("ERROR: "+cont_error);
                                con.end();
                          })
                          
                      });
                  });
            });
          
        }

        if(tipo==="API"){
            let data=[];

          if(producto==="COPPEL"){
              var sql = "SELECT count(mo.estatus) repeticiones, mog.FECHA,mo.email,mog.ID,mo.estatus FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.ID";

              var sql1="SELECT mo2.FECHA,mo1.email,mo2.ID,mo1.estatus from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.ID) group by mo2.ID";

              var sql2="SELECT distinct mo.FECHA,(er.email),mo.ID,er.estatus FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.ID";

              var sql3 = "SELECT count(mo.estatus) repeticiones,mog.FECHA,mo.email,mog.ID,mo.estatus FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.ID";
          }
          
              con.query(sql, function (err, results) {
                  if(results!==undefined){
                      cont_ok=results.length;
                      
                  }
                  open=results;

                  if (err) throw err;
                  con.query(sql1, function(err,results1){
                      if(results1!==undefined){
                          cont_ok=cont_ok+results1.length;
                          
                      }
                      delivered=results1;

                      con.query(sql2, function(err,results2){
                          if(results2!==undefined){
                              cont_error=results2.length;
                              
                          }
                          error=results2;

                          con.query(sql3,function(err,results3){
                              click=results3
                              console.log("crear excel");
                              var newWB = xlsx.utils.book_new()

                              if(ciclo===1){
                                  console.log("se crea el archivo de cero");
                                  
                                    if(open!==undefined){
                                        open.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                  
                                  if(delivered!==undefined){
                                      delivered.forEach(result => {
                                          let dat={};
                                          // console.log(result);
                                          dat.FECHA=result.FECHA;
                                          dat.PRODUCTO=""+producto;
                                          dat.EMAIL=result.email;
                                          dat.IDENTIFICADOR=result.ID;
                                          dat.ESTATUS=result.estatus;
                                          dat.REPETICIONES=result.repeticiones;
                                          dat.VALIDO="1";
                                          dat.INVALIDO="";
                                          data.push(dat);
                                      });
                                  }
                                  
                                    if(error!==undefined){
                                        error.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="";
                                            dat.INVALIDO="1";
                                            data.push(dat);
                                        });
                                    }

                                    if(click!==undefined){
                                        click.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                  
                                  // console.log(data);
                                  var newWS = xlsx.utils.json_to_sheet(data);
                                  xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                  xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
                              }
                              if(ciclo===2){
                                  console.log("se abre archivo y luego se agregan datos");
                                  let data_part=[];

                                  readXlsxFile('./listas/'+producto+'/'+tipo+'/'+mes+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx').then((rows) => {
                  
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
                                      if(open!==undefined){
                                        open.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                  
                                  if(delivered!==undefined){
                                      delivered.forEach(result => {
                                          let dat={};
                                          // console.log(result);
                                          dat.FECHA=result.FECHA;
                                          dat.PRODUCTO=""+producto;
                                          dat.EMAIL=result.email;
                                          dat.IDENTIFICADOR=result.ID;
                                          dat.ESTATUS=result.estatus;
                                          dat.REPETICIONES="";
                                          dat.VALIDO="1";
                                          dat.INVALIDO="";
                                          data.push(dat);
                                      });
                                  }
                                  
                                    if(error!==undefined){
                                        error.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES="";
                                            dat.VALIDO="";
                                            dat.INVALIDO="1";
                                            data.push(dat);
                                        });
                                    }

                                    if(click!==undefined){
                                        click.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=result.FECHA;
                                            dat.PRODUCTO=""+producto;
                                            dat.EMAIL=result.email;
                                            dat.IDENTIFICADOR=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            dat.REPETICIONES=result.repeticiones;
                                            dat.VALIDO="1";
                                            dat.INVALIDO="";
                                            data.push(dat);
                                        });
                                    }
                                      // console.log(data);
                                      var newWS = xlsx.utils.json_to_sheet(data);
                                      xlsx.utils.book_append_sheet(newWB,newWS,"hoja1")//workbook name as param
                                      xlsx.writeFile(newWB,''+direc+'/00MUTA-REPORTE-MAIL-GENERAL-'+mes+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});

                                  })
                              }
                              console.log("EXITOSOS: "+cont_ok);
                              console.log("ERROR: "+cont_error);
                              con.end();
                          })
                          
                      });
                  });
            });
          //}
        }

        

    });//final funcion

}