const { quoteToDoubleQuote } = require("./utils");

const invoice = {
  cabecera: {
    tipOperacion: "0101",
    fecEmision: "2021-06-30",
    horEmision: "10:20:14",
    fecVencimiento: "-",
    codLocalEmisor: "0000",
    tipDocUsuario: "1",
    numDocUsuario: "00000000",
    rznSocialUsuario: "varios",
    tipMoneda: "PEN", // USD EUR
    sumTotTributos: "27.00",
    sumTotValVenta: "150.00",
    sumPrecioVenta: "177.00",
    sumDescTotal: "0.00",
    sumOtrosCargos: "0.00",
    sumTotalAnticipos: "0.00",
    sumImpVenta: "177.00",
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
      mtoValorUnitario: "50.0000000000",
      sumTotTributosItem: "18.00",

      codTriIGV: "1000",
      mtoIgvItem: "18.00",
      mtoBaseIgvItem: "100.00",
      nomTributoIgvItem: "IGV",
      codTipTributoIgvItem: "VAT",
      tipAfeIGV: "10",
      porIgvItem: "18.00",

      mtoPrecioVentaUnitario: "59.00", // (mtoValorVentaItem + mtoIgvItem) / ctdUnidadItem
      mtoValorVentaItem: "100.00", // mtoValorUnitario * ctdUnidadItem
    },
    {
      codUnidadMedida: "NIU",
      ctdUnidadItem: "5.00",
      codProducto: "CD0002",
      codProductoSUNAT: "-",
      desItem: "Producto 2",
      mtoValorUnitario: "10.0000000000",
      sumTotTributosItem: "9.00",

      codTriIGV: "1000",
      mtoIgvItem: "9.00",
      mtoBaseIgvItem: "50.00",
      nomTributoIgvItem: "IGV",
      codTipTributoIgvItem: "VAT",
      tipAfeIGV: "10",
      porIgvItem: "18.00",

      mtoPrecioVentaUnitario: "11.80", // (mtoValorVentaItem + mtoIgvItem) / ctdUnidadItem
      mtoValorVentaItem: "50.00", // mtoValorUnitario * ctdUnidadItem
    },
  ],
  tributos: [
    {
      ideTributo: "1000",
      nomTributo: "IGV",
      codTipTributo: "VAT",
      mtoBaseImponible: "150.00",
      mtoTributo: "27.00"
    },
  ],
  leyendas: [
    {
      codLeyenda: "1000",
      desLeyenda: "CIENTO SETENTA Y SIETE CON 00/100 SOLES",
    }
  ]
}

console.log(JSON.stringify(invoice));

const {cabecera, detalle, tributos, leyendas} = invoice;
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({cabecera})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({detalle})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({tributos})));
console.log();
console.log(quoteToDoubleQuote(JSON.stringify({leyendas})));

