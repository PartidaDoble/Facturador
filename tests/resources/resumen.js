const { quoteToDoubleQuote } = require("./utils");

const resumen = {
  resumenDiario: [
    {
      fecEmision: "2021-07-25",
      fecResumen: "2021-07-26",
      tipDocResumen: "03",
      idDocResumen: "B001-00000001",
      tipDocUsuario: "1",
      numDocUsuario: "00000000",
      tipMoneda: "PEN",
      totValGrabado: "100.00",
      totValExoneado: "0.00",
      totValInafecto: "0.00",
      totValExportado: "0.00",
      totValGratuito: "0.00",
      totOtroCargo: "0.00",
      totImpCpe: "118.00",
      tipEstado: "1",
      tributosDocResumen: [
        {
          idLineaRd: "1",
          ideTributoRd: "1000",
          nomTributoRd: "IGV",
          codTipTributoRd: "VAT",
          mtoBaseImponibleRd: "100.00",
          mtoTributoRd: "18.00",
        }
      ]
    },
    {
      fecEmision: "2021-07-25",
      fecResumen: "2021-07-26",
      tipDocResumen: "03",
      idDocResumen: "B001-00000002",
      tipDocUsuario: "1",
      numDocUsuario: "00000000",
      tipMoneda: "PEN",
      totValGrabado: "200.00",
      totValExoneado: "0.00",
      totValInafecto: "0.00",
      totValExportado: "0.00",
      totValGratuito: "0.00",
      totOtroCargo: "0.00",
      totImpCpe: "236.00",
      tipEstado: "1",
      tributosDocResumen: [
        {
          idLineaRd: "2",
          ideTributoRd: "1000",
          nomTributoRd: "IGV",
          codTipTributoRd: "VAT",
          mtoBaseImponibleRd: "200.00",
          mtoTributoRd: "36.00",
        }
      ]
    }
  ]
}

console.log(JSON.stringify(resumen));
//console.log(quoteToDoubleQuote(JSON.stringify(resumen)));
