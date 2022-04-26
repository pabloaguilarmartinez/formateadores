import Headers from '../constants/HeadersConfigurador';
import Plantillas from 'src/constants/MaestroPlantillasEsamur';
const Excel = require('exceljs');

// Variables globales para recopilar información de procesos y las instancias necesarias para formatear
var arrayProcesos = [];
var arrayInstancias = [];

/**
 * Función que exportamos para obtener los parámetros de abajo y llamar a las demás funciones del servicio
 * @param {*} file Archivo que nos importan
 * @param {string} nombreEdar Nombre de la edar que se usará para el nombre del fichero que se genera y para la dirección de PLC
 * @param {string} identificador Identificador XXX para los tags
 */
export function format(file, nombreEdar, identificador) {
  readFile(file, nombreEdar, identificador);
};

/**
 * Función que comienza leyendo el fichero que nos proporcionan para recoger la información necesaria para formatear
 * @param {*} file
 * @param {string} nombreEdar
 * @param {string} identificador
 */
async function readFile(file, nombreEdar, identificador) {
  const wb = new Excel.Workbook();
  const reader = new FileReader();

  // Por si acaso limpiamos los arrays
  arrayInstancias = [];
  arrayProcesos = [];

  // Leemos fichero y guardamos en los arrays según la funcionalidad deseada
  reader.readAsArrayBuffer(file);
  reader.onload = () => {
    const buffer = reader.result;
    wb.xlsx.load(buffer).then(workbook => {
      const procesosWorksheet = workbook.getWorksheet('Procesos');
      procesosWorksheet.eachRow((row, rowIndex) => {
        if (rowIndex > 1 && row.values[1] !== undefined) {
          arrayProcesos.push({
            rowNumber: rowIndex,
            proceso: row.values[2],
            descripcion: row.values[3],
            numero: parseInt(row.values[4])
          });
        }
      });
      // console.table(arrayProcesos);
      const filtradoAquatecWorksheet = workbook.getWorksheet('Filtrado_Aquatec');
      filtradoAquatecWorksheet.eachRow((row, rowIndex) => {
        if (rowIndex > 1 && row.values[1] !== undefined) {
          arrayInstancias.push({
            automatico: row.values[1].result,
            descripcion: row.values[2],
            agrupacion: row.values[4],
            instancia: row.values[5],
            atributo: row.values[6],
            unidades: row.values[7],
            tag: row.values[15],
            tipoDato: row.values[16],
            direccionPlc: row.values[17],
            grupo: row.values[19],
            estacion: row.values[20],
            revisar: (row.values[6] === 'E_AVER' && row.values[13] !== 'E_AVER') ? true : false,
            minValue: row.values[8],
            maxValue: row.values[9],
            informacionSofrel: row.values[10],
            varCero: null
          });
        }
      });
      // console.table(arrayInstancias);
    }).then(() => {
      // Llamamos a la función que recorre el array de instancais para ver si falta algún atributo que apunte a VAR_0
      checkVarCero();
    }).then(() => {
      // Llamamos a la función que crea el archivo formateado una vez leemos todo
      createFile(nombreEdar, identificador);
    });
  };
};

/**
 * Función que con los datos obtenidos en la función anterior crea la lista de señales formateada
 * @param {string} nombreEdar
 * @param {string} identificador
 */
async function createFile(nombreEdar, identificador) {
  const workbook = new Excel.Workbook();
  // Añadimos una hoja al excel con el nombre Hoja 1
  const worksheet = workbook.addWorksheet('Hoja 1');

  // Asignamos el nombre de los headers que va a tener el nuevo archivo
  worksheet.columns = Headers;

  // Recorremos el array de instancias y lo añadimos al nuevo excel formateando los datos
  arrayInstancias.forEach((instancia) => {
    let proceso = arrayProcesos.find(p => p.proceso === instancia.automatico.substring(0, 4));
    let numero = null;
    if (proceso.numero < 10) {
      numero = '0' + proceso.numero;
    } else {
      numero = proceso.numero;
    }
    worksheet.addRow({
      automatico: identificador + '0100' + numero + 'ED_' + instancia.automatico,
      descartable: (instancia.revisar && instancia.varCero === null) ? 'Yes' : 'No',
      revisar: (instancia.revisar || instancia.instancia.substring(0, 4) === 'ADBA') ? 'Revisar' : '',
      estacion: proceso.numero,
      agrupador: instancia.agrupacion,
      instancia: instancia.instancia,
      atributo: instancia.atributo,
      nombre: instancia.tag,
      tipo: instancia.tipoDato,
      grupo: 'Estación ' + proceso.numero,
      descripcion: instancia.descripcion + ' - ' + proceso.descripcion,
      offMsg: instancia.atributo === 'E_MARC' ? 'Paro' : (instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI') ? 'No' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'Normal' : '',
      onMsg: instancia.atributo === 'E_MARC' ? 'Marcha' : (instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI') ? 'Si' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'Alarma' : '',
      readOnly: 'Yes',
      invertida: (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'Direct' : '',
      engUnits: instancia.unidades === "-" ? '' : instancia.unidades,
      minValue: instancia.minValue,
      maxValue: instancia.maxValue,
      minRaw: instancia.minValue,
      maxRaw: instancia.maxValue,
      historico: (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'No' : 'Yes',
      evento: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? 'Yes' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2)) === 'E_' ? 'No' : '',
      alarmState: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? 'None' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'On' : '',
      alarmPri: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? '' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 400 : '',
      direccionPlc: instancia.informacionSofrel !== null ? 'Sofrel.' + identificador + '.EDAR_' + nombreEdar.toUpperCase() + '.' + ((instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'LI_' : 'NI_') + (instancia.informacionSofrel < 10 ? '000' + instancia.informacionSofrel : instancia.informacionSofrel < 100 ? '00' + instancia.informacionSofrel : instancia.informacionSofrel < 1000 ? '0' + instancia.informacionSofrel : instancia.informacionSofrel) + '.Value' : '',
      varCero: instancia.varCero
    });
  });

  // Añadimos la instancias de Comunicación Sofrel que siempre se añade
  // Ordenamos el array de procesos para que el número de estación sea el mayor + 1
  arrayProcesos.sort((a, b) => {
    if (a.numero == b.numero) {
      return 0;
    }
    if (a.numero < b.numero) {
      return -1;
    }
    if (a.numero > b.numero) {
      return 1;
    }
  });
  const proceso = arrayProcesos[arrayProcesos.length - 1];
  let numero = null;
  if (proceso.numero < 10) {
    numero = '0' + (proceso.numero + 1);
  } else {
    numero = (proceso.numero + 1);
  }
  worksheet.addRows([{
    automatico: identificador + '0100' + numero + 'ED_SCBA01.E_ESTA',
    descartable: 'No',
    revisar: 'Revisar, atributo añadido automáticamente',
    estacion: proceso.numero,
    descripcion: 'Comunicación Establecida',
    instancia: 'SCBA01',
    atributo: 'E_ESTA',
    tipo: 'DIGITAL',
    invertida: 'Direct',
    historico: 'Yes',
    evento: 'Yes',
    alarmState: 'None',
    direccionPlc: 'Sofrel.' + identificador + '.EDAR_' + nombreEdar.toUpperCase() + '.Connection.Established',
    grupo: 'Estación ' + proceso.numero
  }, {
    automatico: identificador + '0100' + numero + 'ED_SCBA01.E_FCOM',
    descartable: 'No',
    revisar: 'Revisar, atributo añadido automáticamente',
    estacion: proceso.numero,
    descripcion: 'Fallo Comunicaciones',
    instancia: 'SCBA01',
    atributo: 'E_FCOM',
    tipo: 'DIGITAL',
    invertida: 'Direct',
    historico: 'Yes',
    evento: 'Yes',
    alarmState: 'None',
    grupo: 'Estación ' + proceso.numero,
    direccionPlc: 'Sofrel.' + identificador + '.EDAR_' + nombreEdar.toUpperCase() + '.Connection.Failed'
  },
  {
    automatico: identificador + '0100' + numero + 'ED_SCBA01.T_ASKD',
    descartable: 'No',
    revisar: 'Revisar, atributo añadido automáticamente',
    estacion: proceso.numero,
    descripcion: 'Telemando Petición Datos',
    instancia: 'SCBA01',
    atributo: 'T_ASKD',
    tipo: 'DIGITAL',
    invertida: 'Direct',
    historico: 'Yes',
    evento: 'Yes',
    alarmState: 'None',
    grupo: 'Estación ' + proceso.numero,
    direccionPlc: 'Sofrel.' + identificador + '.EDAR_' + nombreEdar.toUpperCase() + '.Connection.AskedPrimary'
  }]);

  // Guardamos el archivo
  await workbook.xlsx.writeBuffer({
    based64: true
  }).then((xls64) => {
    var a = document.createElement("a");
    var data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(data);
    a.href = url;
    const today = new Date();
    a.download = 'ESAMUR - LS EDAR ' + nombreEdar + ' ' + (today.getDate() < 10 ? ('0' + today.getDate()) : today.getDate()) + (today.getMonth() < 9 ? ('0' + (today.getMonth() + 1)) : (today.getMonth() + 1)) + today.getFullYear() + '.xlsx';
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

async function checkVarCero() {
  await arrayInstancias.forEach((instancia) => {
    let atributos = Plantillas[instancia.instancia.substring(0, 4)];
    atributos.forEach(atributo => {
      if (arrayInstancias.findIndex(i => atributo === i.atributo && i.instancia === instancia.instancia && i.agrupacion === instancia.agrupacion) === -1) {
        arrayInstancias.push({
          automatico: instancia.automatico.substring(0, 12) + atributo,
          descripcion: null,
          agrupacion: instancia.agrupacion,
          instancia: instancia.instancia,
          atributo: atributo,
          unidades: null,
          tag: null,
          tipoDato: null,
          direccionPlc: null,
          grupo: instancia.grupo,
          estacion: instancia.estacion,
          revisar: true,
          minValue: null,
          maxValue: null,
          informacionSofrel: null,
          varCero: 'VAR_0'
        });
      }
    });
  });
};
