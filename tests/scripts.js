const boleta = {
  "cabecera": {
    tipOperacion: "0101",
    fecEmision: "2021-06-28",
    horEmision: "10:20:14",
    fecVencimiento: "-", // que pasa si se omite?
    codLocalEmisor: "0000",
    tipDocUsuario: "0",
    numDocUsuario: "00000000",
    rznSocialUsuario: "varios",
    tipMoneda: "PEN", // USD EUR
    sumTotTributos: "18.00",
    sumTotValVenta: "100.00",
    sumPrecioVenta: "118.00",
    sumDescTotal: "0.00",
    sumOtrosCargos: "0.00",
    sumTotalAnticipos: "0.00",
    sumImpVenta: "118.00", // que pasa si omito los decimales?
    ublVersionId: "2.1",
    customizationId: "2.0",
  },
  "detalle": [
    {
      codUnidadMedida: "NIU",
      ctdUnidadItem: "2.00",
      codProducto: "CD0001",
      codProductoSUNAT: "-",
      desItem: "Producto 1",
      mtoValorUnitario: "50.00",
      sumTotTributosItem: "18.00",

      codTriIGV: "1000",
      mtoIgvItem: "18.00",
      mtoBaseIgvItem: "100.00",
      nomTributoIgvItem: "IGV",
      codTipTributoIgvItem: "VAT",
      tipAfeIGV: "10",
      porIgvItem: "18.0",

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
}

console.log(JSON.stringify(boleta));

//console.log(quoteToDoubleQuote(JSON.stringify(boleta)));

function quoteToDoubleQuote(string) {
  return string.replace(/"/g, '""');
}
