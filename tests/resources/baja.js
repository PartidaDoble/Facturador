const { quoteToDoubleQuote } = require("./utils");

const bajas = {
  resumenBajas: [
    {
      fecGeneracion: "2021-07-18",
      fecComunicacion: "2021-07-18",
      tipDocBaja: "01",
      numDocBaja: "F001-00000007",
      desMotivoBaja: "ERROR EN EL NOMBRE DEL CLIENTE"
    },
    {
      fecGeneracion: "2021-07-18",
      fecComunicacion: "2021-07-18",
      tipDocBaja: "01",
      numDocBaja: "F001-00000008",
      desMotivoBaja: "ERROR EN EL NUMERO DE RUC"
    }
  ]
}

console.log(JSON.stringify(bajas));

// const {cabecera, detalle, tributos, leyendas} = invoice;
// console.log();
console.log(quoteToDoubleQuote(JSON.stringify({bajas})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({detalle})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({tributos})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({leyendas})));

