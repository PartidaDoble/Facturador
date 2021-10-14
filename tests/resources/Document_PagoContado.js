const { quoteToDoubleQuote } = require("./utils");

const invoice = {
  cabecera: {
    tipOperacion: "0101",
    fecEmision: "2021-08-31",
    horEmision: "10:20:14",
    fecVencimiento: "-",
    codLocalEmisor: "0000",
    tipDocUsuario: "6",
    numDocUsuario: "20131380951",
    rznSocialUsuario: "CLIENTE SAC",
    tipMoneda: "PEN",
    sumTotTributos: "18.00",
    sumTotValVenta: "100.00",
    sumPrecioVenta: "118.00",
    sumDescTotal: "0.00",
    sumOtrosCargos: "0.00",
    sumTotalAnticipos: "0.00",
    sumImpVenta: "118.00",
    ublVersionId: "2.1",
    customizationId: "2.0",
  },
  detalle: [
    {
      codUnidadMedida: "NIU",
      ctdUnidadItem: "2.00",
      codProducto: "10000",
      codProductoSUNAT: "-",
      desItem: "Producto 1",
      mtoValorUnitario: "50.00000000",
      sumTotTributosItem: "18.00",

      codTriIGV: "1000",
      mtoIgvItem: "18.00",
      mtoBaseIgvItem: "100.00",
      nomTributoIgvItem: "IGV",
      codTipTributoIgvItem: "VAT",
      tipAfeIGV: "10",
      porIgvItem: "18.00",

      mtoPrecioVentaUnitario: "59.00",
      mtoValorVentaItem: "100.00",
    },
  ],
  tributos: [
    {
      ideTributo: "1000",
      nomTributo: "IGV",
      codTipTributo: "VAT",
      mtoBaseImponible: "100.00",
      mtoTributo: "18.00",
    },
  ],
  leyendas: [
    {
      codLeyenda: "1000",
      desLeyenda: "CIENTO DIECIOCHO CON 00/100 SOLES",
    }
  ],
  datoPago: {
    formaPago: "Contado",
    mtoNetoPendientePago: "0.00",
    tipMonedaMtoNetoPendientePago: "PEN"
  }
}

console.log(JSON.stringify(invoice));

const {cabecera, detalle, tributos, leyendas, datoPago} = invoice;
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({cabecera})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({detalle})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({tributos})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({leyendas})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({datoPago})));
