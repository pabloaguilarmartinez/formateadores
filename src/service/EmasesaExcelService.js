import Headers from '../constants/HeadersListaEmasesa';
import Activos from 'src/constants/ActivosTipoElementoEmasesa';
import Unidades from 'src/constants/UnidadesEmasesa';
const Excel = require('exceljs');

// Variables globales para recopilar información
var arraySignals = [];
var arrayInfoCodificacion = [];

/**
 * Función que exportamos para coger el fichero importado y proceder a leer y crear el nuevo fichero con las señales asociadas
 * @param {*} file Archivo que nos importan
 */
export function associate(file, scer, scaa) {
  readFile(file, scer, scaa);
};

/**
 * Función para leer la información necesaria del fichero
 * @param {*} file
 */
async function readFile(file, scer, scaa) {
  const wb = new Excel.Workbook();
  const reader = new FileReader();

  // Limpiamos arrays
  arrayInfoCodificacion = [];
  arraySignals = [];

  // Leemos fichero y desdoblamos
  reader.readAsArrayBuffer(file);
  reader.onload = () => {
    const buffer = reader.result;
    wb.xlsx.load(buffer).then(workbook => {
      const signalsWorksheet = workbook.getWorksheet(1);
      // const infoCodificacion = workbook.getWorksheet('INFO CODIFICACION');
      signalsWorksheet.eachRow((row, rowIndex) => {
        if (rowIndex > 1 && row.values[1] !== undefined) {
          arraySignals.push({
            name: row.values[1],
            rtu: row.values[2],
            estacion: row.values[3],
            type: row.values[4],
            description: row.values[5],
            iengrRawmin: row.values[6],
            iengrRawmax: row.values[7],
            iengrEgumin: row.values[8],
            iengrEgumax: row.values[9],
            iengrScaleraw: row.values[10],
            units: row.values[11],
            hilowDoit: row.values[12],
            hilowDoitdoit: row.values[13],
            hilowDead: row.values[14],
            hilowLolim: row.values[15],
            hilowHilim: row.values[16],
            hilowLololim: row.values[17],
            hilowHihilim: row.values[18],
            anainType: row.values[19],
            scale: row.values[20],
            abnrmStateAlmZero: row.values[21],
            abnrmStateAlmOne: row.values[22],
            flagBmsg: row.values[23],
            outputCmdZero: row.values[24],
            outputCmdTwo: row.values[25],
            outputCmdOne: row.values[26],
            outputCmdThree: row.values[27],
            anainIospecExternal: row.values[28],
            inbitBitdefZeroIospecExternal: row.values[29],
            inbitBitdefOneIospecExternal: row.values[30],
            accinIospecExternal: row.values[31],
            anainIospecExternalTwo: row.values[32],
            input: row.values[33],
            anaoutIospecExternal: row.values[34],
            outsOneIospecExternal: row.values[35],
            outsTwoIospecExternal: row.values[36],
            output: row.values[37],
            sustainCosAlarm: row.values[38],
            elemento: row.values[39],
            tipoElemento: row.values[40],
            cotoutActivo: row.values[41],
            hilowDinamica: row.values[42],
            supressionType: row.values[43],
            parentStrKey: row.values[44],
            timeout: row.values[45],
            rtnAlmTimeout: row.values[46],
            holdOffTimeout: row.values[47],
            flagAlminh: row.values[48],
            flagClrinh: row.values[49],
            alarma: row.values[50],
            flagEvtin: row.values[51],
            flagCevin: row.values[52],
            evento: row.values[53],
            historico: row.values[54],
            // revisionColumnaFinal: row.values[68].result
          });
        }
      });
    })
    .then(() => {
      createFile(scer,scaa);
    });
  };
};

/**
 * Función para crear un nuevo fichero con el mismo formato que la hoja de estación de EMASESA
 * y no tocar el original por si se quisiera revisar
 */
async function createFile(scer, scaa) {
  const workbook = new Excel.Workbook();
  // Añadimos una hoja al excel con el nombre Hoja 1
  const worksheet = workbook.addWorksheet('Hoja 1');

  // Asignamos el nombre de los headers que va a tener el nuevo archivo
  worksheet.columns = Headers;

  // Recorremos array de señales y lo añadimos al nuevo excel
  // Desdoblando, adecuando y rellenando lo que falta y se puede automatizar
  arraySignals.forEach(signal => {
    let signalToBeUnfolded = false;
    if (signal.inbitBitdefZeroIospecExternal !== undefined && signal.inbitBitdefOneIospecExternal !== undefined) {
      signalToBeUnfolded = true;
    }

    worksheet.addRow({
      name: signal.name,
      rtu: signal.rtu,
      estacion: signal.estacion,
      type: signalToBeUnfolded ? 'ANALOG' : signal.type,
      description: getCleanedString(signal.description),
      iengrRawmin: signal.iengrRawmin,
      iengrRawmax: signal.iengrRawmax,
      iengrEgumin: signal.iengrEgumin,
      iengrEgumax: signal.iengrEgumax,
      iengrScaleraw: signal.iengrScaleraw,
      units: signalToBeUnfolded ? 'ud' : (signal.units in Unidades) ? Unidades[signal.units] : signal.units,
      hilowDoit: signal.hilowDoit,
      hilowDoitdoit: signal.hilowDoitdoit,
      hilowDead: signal.hilowDead,
      hilowLolim: signal.hilowLolim,
      hilowHilim: signal.hilowHilim,
      hilowLololim: signal.hilowLololim,
      hilowHihilim: signal.hilowHihilim,
      anainType: signal.anainType,
      scale: signal.scale,
      abnrmStateAlmZero: signal.abnrmStateAlmZero,
      abnrmStateAlmOne: signal.abnrmStateAlmOne,
      flagBmsg: signal.flagBmsg,
      outputCmdZero: signal.outputCmdZero,
      outputCmdTwo: signal.outputCmdTwo,
      outputCmdOne: signal.outputCmdOne,
      outputCmdThree: signal.outputCmdThree,
      anainIospecExternal: signal.anainIospecExternal,
      inbitBitdefZeroIospecExternal: signal.inbitBitdefZeroIospecExternal,
      inbitBitdefOneIospecExternal: signal.inbitBitdefOneIospecExternal,
      accinIospecExternal: signal.accinIospecExternal,
      anainIospecExternalTwo: signal.anainIospecExternalTwo,
      input: (signal.anainIospecExternal !== '' && signal.anainIospecExternal !== undefined) ? signal.anainIospecExternal : (signal.inbitBitdefOneIospecExternal !== '' && signal.inbitBitdefOneIospecExternal !== undefined && signal.inbitBitdefZeroIospecExternal !== '' && signal.inbitBitdefZeroIospecExternal !== undefined) ? signal.inbitBitdefZeroIospecExternal : (signal.accinIospecExternal !== '' && signal.accinIospecExternal !== undefined) ? signal.accinIospecExternal : (signal.anainIospecExternalTwo !== '' && signal.anainIospecExternalTwo !== undefined) ? signal.anainIospecExternalTwo : '',
      anaoutIospecExternal: signal.anaoutIospecExternal,
      outsOneIospecExternal: signal.outsOneIospecExternal,
      outsTwoIospecExternal: signal.outsTwoIospecExternal,
      output: (signal.anaoutIospecExternal !== '' && signal.anaoutIospecExternal !== undefined) ? signal.anaoutIospecExternal : (signal.outsOneIospecExternal !== '' && signal.outsOneIospecExternal !== undefined && signal.outsTwoIospecExternal !== '' && signal.outsTwoIospecExternal !== undefined) ? signal.outsOneIospecExternal : '',
      sustainCosAlarm: signal.sustainCosAlarm,
      elemento: signal.elemento,
      tipoElemento: signal.tipoElemento,
      cotoutActivo: signal.cotoutActivo,
      hilowDinamica: signal.hilowDinamica,
      supressionType: signal.supressionType,
      parentStrKey: signal.parentStrKey,
      timeout: signal.timeout,
      rtnAlmTimeout: signal.rtnAlmTimeout,
      holdOffTimeout: signal.holdOffTimeout,
      flagAlminh: signal.flagAlminh,
      flagClrinh: signal.flagClrinh,
      alarma: (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'no' : signalToBeUnfolded ? 'yes' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'no' : 'yes',
      flagEvtin: signal.flagEvtin,
      flagCevin: signal.flagCevin,
      evento: (signal.name.substring(0, 2) == 'XE' && (signal.name.substring(signal.name.length - 2, signal.name.length) == 'EN' || signal.name.substring(signal.name.length - 2, signal.name.length) == 'SI' || signal.name.substring(signal.name.length - 2, signal.name.length) == 'ED' || signal.name.substring(signal.name.length - 2, signal.name.length) == 'RS')) ? 'yes' : (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'yes' : signalToBeUnfolded ? 'no' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'yes' : 'no',
      historico: (signal.name.substring(0, 2) == 'ER') ? signal.historico : (signal.name.substring(0, 2) == 'XE' && (signal.name.substring(signal.name.length - 2, signal.name.length) == 'FR' || signal.name.substring(signal.name.length - 2, signal.name.length) == 'FW' || signal.name.substring(signal.name.length - 2, signal.name.length) == 'TX')) ? 'yes' : (signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '1' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '1' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '2' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '3' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '4' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === '6' || signal.name.substring(signal.estacion.split("_")[0].length, signal.estacion.split("_")[0].length + 1) === 'A') ? 'yes' : 'no',
      revisionColumnaFinal: signal.revisionColumnaFinal,
      scActivo: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scActivo + '01' : '',
      scServicio: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scServicio : '',
      scProceso: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scProceso : '',
      scInstalacion: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scInstalacion : '',
      scAtributo: getAtributo(signal),
      scer: scer + '',
      scaa: scaa
    });

    // Si hay que desdoblar la señal se añaden las dos nuevas
    if (signalToBeUnfolded) {
      // .1
      worksheet.addRow({
        name: signal.name + '.1',
        rtu: signal.rtu,
        estacion: signal.estacion,
        type: signal.type,
        description: getCleanedString(signal.description + ' ' + signal.outputCmdOne),
        iengrRawmin: signal.iengrRawmin,
        iengrRawmax: signal.iengrRawmax,
        iengrEgumin: signal.iengrEgumin,
        iengrEgumax: signal.iengrEgumax,
        iengrScaleraw: signal.iengrScaleraw,
        units: signal.units,
        hilowDoit: signal.hilowDoit,
        hilowDoitdoit: signal.hilowDoitdoit,
        hilowDead: signal.hilowDead,
        hilowLolim: signal.hilowLolim,
        hilowHilim: signal.hilowHilim,
        hilowLololim: signal.hilowLololim,
        hilowHihilim: signal.hilowHihilim,
        anainType: signal.anainType,
        scale: signal.scale,
        abnrmStateAlmZero: signal.abnrmStateAlmZero,
        abnrmStateAlmOne: signal.abnrmStateAlmOne,
        flagBmsg: signal.flagBmsg,
        outputCmdZero: signal.outputCmdZero,
        outputCmdTwo: signal.outputCmdOne === 'AUTOMÁTICO' || signal.outputCmdOne === 'MANUAL' || signal.outputCmdOne === 'AUTOMATICO' || signal.outputCmdOne === 'HABILITADO' || signal.outputCmdOne === 'ACTIVO' || signal.outputCmdOne === 'EN SERVICIO' || signal.outputCmdOne === 'LOCAL' ? 'NO ESTADO' : 'NO ' + signal.outputCmdOne,
        outputCmdOne: signal.outputCmdOne,
        outputCmdThree: signal.outputCmdThree,
        anainIospecExternal: signal.anainIospecExternal,
        inbitBitdefZeroIospecExternal: signal.inbitBitdefZeroIospecExternal,
        accinIospecExternal: signal.accinIospecExternal,
        anainIospecExternalTwo: signal.anainIospecExternalTwo,
        input: signal.inbitBitdefZeroIospecExternal,
        anaoutIospecExternal: signal.anaoutIospecExternal,
        outsOneIospecExternal: signal.outsOneIospecExternal,
        outsTwoIospecExternal: signal.outsTwoIospecExternal,
        output: signal.output,
        sustainCosAlarm: signal.sustainCosAlarm,
        elemento: signal.elemento,
        tipoElemento: signal.tipoElemento,
        cotoutActivo: signal.cotoutActivo,
        hilowDinamica: signal.hilowDinamica,
        supressionType: signal.supressionType,
        parentStrKey: signal.parentStrKey,
        timeout: signal.timeout,
        rtnAlmTimeout: signal.rtnAlmTimeout,
        holdOffTimeout: signal.holdOffTimeout,
        flagAlminh: signal.flagAlminh,
        flagClrinh: signal.flagClrinh,
        alarma: (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'no' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'no' : 'yes',
        flagEvtin: signal.flagEvtin,
        flagCevin: signal.flagCevin,
        evento: (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'yes' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'yes' : 'no',
        historico: 'no',
        revisionColumnaFinal: signal.revisionColumnaFinal,
        scActivo: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scActivo + '01' : '',
        scServicio: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scServicio : '',
        scProceso: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scProceso : '',
        scInstalacion: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scInstalacion : '',
        scAtributo: getAtributo(signal),
        scer: scer,
        scaa: scaa
      });
      // .2
      worksheet.addRow({
        name: signal.name + '.2',
        rtu: signal.rtu,
        estacion: signal.estacion,
        type: signal.type,
        description: getCleanedString(signal.description + ' ' + signal.outputCmdTwo),
        iengrRawmin: signal.iengrRawmin,
        iengrRawmax: signal.iengrRawmax,
        iengrEgumin: signal.iengrEgumin,
        iengrEgumax: signal.iengrEgumax,
        iengrScaleraw: signal.iengrScaleraw,
        units: signal.units,
        hilowDoit: signal.hilowDoit,
        hilowDoitdoit: signal.hilowDoitdoit,
        hilowDead: signal.hilowDead,
        hilowLolim: signal.hilowLolim,
        hilowHilim: signal.hilowHilim,
        hilowLololim: signal.hilowLololim,
        hilowHihilim: signal.hilowHihilim,
        anainType: signal.anainType,
        scale: signal.scale,
        abnrmStateAlmZero: signal.abnrmStateAlmZero,
        abnrmStateAlmOne: signal.abnrmStateAlmOne,
        flagBmsg: signal.flagBmsg,
        outputCmdZero: signal.outputCmdZero,
        outputCmdTwo: signal.outputCmdOne === 'AUTOMÁTICO' || signal.outputCmdOne === 'MANUAL' || signal.outputCmdOne === 'AUTOMATICO' || signal.outputCmdOne === 'HABILITADO' || signal.outputCmdOne === 'ACTIVO' || signal.outputCmdOne === 'EN SERVICIO' || signal.outputCmdOne === 'LOCAL' ? 'NO ESTADO' : 'NO ' + signal.outputCmdTwo,
        outputCmdOne: signal.outputCmdTwo,
        outputCmdThree: signal.outputCmdThree,
        anainIospecExternal: signal.anainIospecExternal,
        inbitBitdefZeroIospecExternal: signal.inbitBitdefOneIospecExternal,
        accinIospecExternal: signal.accinIospecExternal,
        anainIospecExternalTwo: signal.anainIospecExternalTwo,
        input: signal.inbitBitdefOneIospecExternal,
        anaoutIospecExternal: signal.anaoutIospecExternal,
        outsOneIospecExternal: signal.outsOneIospecExternal,
        outsTwoIospecExternal: signal.outsTwoIospecExternal,
        output: signal.output,
        sustainCosAlarm: signal.sustainCosAlarm,
        elemento: signal.elemento,
        tipoElemento: signal.tipoElemento,
        cotoutActivo: signal.cotoutActivo,
        hilowDinamica: signal.hilowDinamica,
        supressionType: signal.supressionType,
        parentStrKey: signal.parentStrKey,
        timeout: signal.timeout,
        rtnAlmTimeout: signal.rtnAlmTimeout,
        holdOffTimeout: signal.holdOffTimeout,
        flagAlminh: signal.flagAlminh,
        flagClrinh: signal.flagClrinh,
        alarma: (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'no' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'no' : 'yes',
        flagEvtin: signal.flagEvtin,
        flagCevin: signal.flagCevin,
        evento: (signal.flagAlminh === 'yes' && signal.flagClrinh === 'yes') ? 'yes' : (signal.type === 'DIGITAL' && signal.abnrmStateAlmOne === 'no' && signal.abnrmStateAlmZero === 'no') ? 'yes' : 'no',
        historico: 'no',
        revisionColumnaFinal: signal.revisionColumnaFinal,
        scActivo: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scActivo + '01' : '',
        scServicio: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scServicio : '',
        scProceso: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scProceso : '',
        scInstalacion: (signal.tipoElemento in Activos) ? Activos[signal.tipoElemento].scInstalacion : '',
        scAtributo: getAtributo(signal),
        scer: scer,
        scaa: scaa
      });
    }
  });

  // Guardamos el archivo
  await workbook.xlsx.writeBuffer({
    based64: true
  }).then((xls64) => {
    var a = document.createElement("a");
    var data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(data);
    a.href = url;
    a.download = 'Prueba.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    },
      0);
  }).catch(function (error) {
    console.log(error.message);
  });
};

// Funciones auxiliares
function getCleanedString(text) {
  return text.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
};

function getAtributo(signal) {
  return signal.tipoElemento === 'CARGA' ? 'M_CARG'
    : signal.tipoElemento === 'DISPONIBILIDAD' ? 'M_DISP'
      : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'ED') ? 'E_DIAG'
        : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'EN') ? 'E_REDA'
          : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'FR') ? 'V_FLEC'
            : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'FW') ? 'V_FESC'
              : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'PR') ? 'V_PLEC'
                : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'RS') ? 'T_RSET'
                  : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'SI') ? 'E_SIMU'
                    : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'SR') ? 'V_LCOR'
                      : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'SW') ? 'V_ECOR'
                        : (signal.name.substring(0, 2) == 'XE' && signal.name.substring(signal.name.length - 2, signal.name.length) == 'TX') ? 'V_TXBY'
                          : '';
};
