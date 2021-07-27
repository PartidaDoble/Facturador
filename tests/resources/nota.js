const { quoteToDoubleQuote } = require("./utils");

const notaCredito = {
  cabecera: {
    tipOperacion: "0101",
    fecEmision: "2021-07-25",
    horEmision: "13:01:30",
    codLocalEmisor: "0000",
    tipDocUsuario: "6",
    numDocUsuario: "20448177484",
    rznSocialUsuario: "TEST SAC",
    tipMoneda: "PEN",
    codMotivo: "07",
    desMotivo: "DEVOLUCIÓN DE MERCADERIAS",
    tipDocAfectado: "01",
    numDocAfectado: "F001-00000001",
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
      codProducto: "CD0001",
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

      mtoPrecioVentaUnitario: "59.00", // (mtoValorVentaItem + mtoIgvItem ) / ctdUnidadItem
      mtoValorVentaItem: "100.00", // mtoValorUnitario * ctdUnidadItem
    },
  ],
  tributos: [
    {
      ideTributo: "1000",
      nomTributo: "IGV",
      codTipTributo: "VAT",
      mtoBaseImponible: "100.00",
      mtoTributo: "18.00"
    },
  ],
  leyendas: [
    {
      codLeyenda: "1000",
      desLeyenda: "CIENTO DIECIOCHO CON 00/100 SOLES",
    }
  ]
}

console.log(JSON.stringify(notaCredito));

// const {cabecera, detalle, tributos, leyendas} = invoice;
// console.log();
//console.log(quoteToDoubleQuote(JSON.stringify({bajas})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({detalle})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({tributos})));
// console.log();
// console.log(quoteToDoubleQuote(JSON.stringify({leyendas})));

