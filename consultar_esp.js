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
let subp="";
let productof;

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
      productof=prod;
      
      if(producto==="CARTA_INDEX"){
        readline.question(`SUBPRODUCTO: `, sub => {
            subp=sub;
            console.log(subp);
            productof=producto+"_"+subp;
            return callback(null,"Final entrada");
        })
      }
      else{
        productof=prod;
        return callback(null,"Final entrada");
      }
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
    let direc=Path.join(__dirname+'/listas/'+tipo+'/'+producto+'/'+mes+'');
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

        if(tipo==="ESPECIAL"){
            console.log("se va a generar reporte de caratula");
            let data=[];

            if(producto==="BANORTE_TMK_INB" || producto==="BANORTE_TMK"){
               
                //USAR APARTIR DE LA REMESA DE TMK-INBOUND
                var sql="SELECT count(mo.estatus) repeticiones,mog.REMESA,mo.email,mog.NO_CERTIFICADO,mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO";

                var sql1="SELECT mo2.REMESA,mo1.email, mo2.NO_CERTIFICADO,mo2.FECHA_INICIO,mo2.FECHA_FIN,mo1.fecha_estatus,mo1.estatus, mo2.NOMBRE from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.NO_CERTIFICADO) and mo1.estatus='delivered' group by mo2.NO_CERTIFICADO";

                var sql2="SELECT distinct mo.REMESA,(er.email),mo.NO_CERTIFICADO,mo.FECHA_INICIO,mo.FECHA_FIN,er.fecha_estatus,er.estatus, mo.NOMBRE FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by er.email";

                var sql3="SELECT count(mo.estatus) repeticiones, mog.REMESA,mo.email,mog.NO_CERTIFICADO, mog.FECHA_INICIO,mog.FECHA_FIN,mo.fecha_estatus,mo.estatus, mog.NOMBRE FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.NO_CERTIFICADO"

            }

            if(producto==="CARTA_INDEX"){

               
                //USAR APARTIR DE LA REMESA DE TMK-INBOUND
                var sql="SELECT mo.email,mog.poliza_2 NO_CERTIFICADO,mo.fecha_estatus,mo.estatus FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza_2";

                var sql1="SELECT mo1.email,mo2.poliza_2 NO_CERTIFICADO ,mo1.fecha_estatus,mo1.estatus from cardif."+nombre2+" mo1,cardif."+nombre1+" mo2 where mo1.email=mo2.correo and mo1.email not in(SELECT mo.email FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='open' group by mog.poliza_2) and mo1.estatus='delivered' group by mo2.poliza_2";

                var sql2="SELECT distinct (er.email),mo.poliza_2 NO_CERTIFICADO,er.fecha_estatus,er.estatus FROM cardif."+nombre1+" mo, cardif."+nombre3+" er where mo.correo=er.email group by mo.poliza_2";

                var sql3="SELECT mo.email,mog.poliza_2 NO_CERTIFICADO,mo.fecha_estatus,mo.estatus FROM cardif."+nombre1+" mog, cardif."+nombre2+" mo where mog.correo=mo.email and mo.estatus='click' group by mog.poliza_2"

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

                                //if(ciclo===1){
                                    console.log("se crea el archivo de cero");
                                    
                                    if(open!==undefined){
                                        open.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+productof;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:0;
                                            dat.EMAIL=result.email;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="CARTA_INDEX"?dat.POLIZA=result.NO_CERTIFICADO: dat.POLIZA=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(delivered!==undefined){
                                        delivered.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+productof;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:0;
                                            dat.EMAIL=result.email;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="CARTA_INDEX"?dat.POLIZA=result.NO_CERTIFICADO: dat.POLIZA=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            data.push(dat);
                                        });
                                    }
                                    
                                    if(error!==undefined){
                                        error.forEach(result => {
                                            let dat={};
                                            // console.log(result);
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+productof;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:0;
                                            dat.EMAIL=result.email;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="CARTA_INDEX"?dat.POLIZA=result.NO_CERTIFICADO: dat.POLIZA=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            data.push(dat);
                                        });
                                    }

                                    if(click!==undefined){
                                        click.forEach(result => {
                                            let dat={};
                                            dat.FECHA=""+fecha;
                                            dat.PRODUCTO=""+productof;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB" || producto==="PLENITUD"?dat.REMESA=result.REMESA:0;
                                            dat.EMAIL=result.email;
                                            producto==="BANORTE_TMK" || producto==="BANORTE_TMK_INB"?dat.POLIZA=result.NO_CERTIFICADO:producto==="CARTA_INDEX"?dat.POLIZA=result.NO_CERTIFICADO: dat.POLIZA=result.ID;
                                            dat.ESTATUS=result.estatus;
                                            data.push(dat);
                                        });
                                    }
                                    
                                    // console.log(data);
                                    var total=cont_ok+cont_error;
                                    var ws_data = [
                                        [ "PRODUCTO", "VALIDO", "INVALIDO", "TOTAL"],
                                        [  productof ,  cont_ok ,  cont_error ,  total]
                                    ];
                                    var ws = xlsx.utils.aoa_to_sheet(ws_data);
                                    var newWS = xlsx.utils.json_to_sheet(data);

                                    /* Add the worksheet to the workbook */
                                    xlsx.utils.book_append_sheet(newWB, ws, "GENERAL");
                                    xlsx.utils.book_append_sheet(newWB,newWS,"DETALLE");
                                    xlsx.writeFile(newWB,''+direc+'/00MUTA-'+nombre1+'.xlsx',{Props:{Author:"WESEND"},bookType:'xlsx',bookSST:true});
                                //}
                                console.log("EXITOSOS: "+cont_ok);
                                console.log("ERROR: "+cont_error);
                                console.log("TOTAL: "+total);
                                con.end();
                          })
                          
                      });
                  });
            });
          
        }
    });//final funcion

}