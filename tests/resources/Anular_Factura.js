const { quoteToDoubleQuote } = require("./utils");

const bajas = {
  resumenBajas: [
    {
      fecGeneracion: "2021-08-30",
      fecComunicacion: "2021-08-31",
      tipDocBaja: "01",
      numDocBaja: "F001-00000001",
      desMotivoBaja: "ERROR EN EL RUC"
    },
  ]
}

console.log(JSON.stringify(bajas));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify(bajas)));
