import Headers from '../constants/Headers';
const Excel = require('exceljs');

var arrayProcesos = [];
var arrayInstancias = [];

export function format(file, nombreEdar, identificador) {
  readFile(file, nombreEdar, identificador)
};

async function readFile(file, nombreEdar, identificador) {
  const wb = new Excel.Workbook();
  const reader = new FileReader();
  arrayInstancias = [];
  arrayProcesos = [];

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
            numero: row.values[4]
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
            informacionSofrel: row.values[10]
          });
        }
      });
      // console.table(arrayInstancias);
    }).then(() => {
      createFile(nombreEdar, identificador);
    });
  };
};

async function createFile(nombreEdar, identificador) {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Hoja 1');

  worksheet.columns = Headers;

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
      descartable: instancia.revisar ? 'Yes' : 'No',
      revisar: (instancia.revisar || instancia.instancia.substring(0, 4) === 'ADBA') ? 'Revisar' : '',
      estacion: proceso.numero,
      agrupador: instancia.agrupacion,
      instancia: instancia.instancia,
      atributo: instancia.atributo,
      nombre: instancia.tag,
      tipo: instancia.tipoDato,
      grupo: 'Estación ' + proceso.numero,
      descripcion: instancia.descripcion,
      offMsg: instancia.atributo === 'E_MARC' ? 'Paro' : (instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI') ? 'No' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'Normal' : '',
      onMsg: instancia.atributo === 'E_MARC' ? 'Marcha' : (instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI') ? 'Si' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'Alarma' : '',
      readOnly: 'Yes',
      invertida: (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ?  'Direct' : '',
      engUnits: instancia.unidades === "-" ? '' : instancia.unidades,
      minValue: instancia.minValue,
      maxValue: instancia.maxValue,
      minRaw: instancia.minValue,
      maxRaw: instancia.maxValue,
      historico: (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'No' : 'Yes',
      evento: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? 'Yes' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2)) === 'E_' ? 'No' : '',
      alarmState: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? 'None' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'On' : '',
      alarmPri: (instancia.atributo === 'E_MARC' || instancia.atributo === 'E_ABIE' || instancia.atributo === 'E_CERR' || instancia.atributo.substring(0, 4) === 'E_DI' || instancia.instancia.substring(0, 4) === 'ADBA') ? '' : (instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 400 : '',
      direccionPlc: 'Sofrel.' + identificador + '.EDAR_' + nombreEdar.toUpperCase() + '.' + ((instancia.tipoDato === 'DIGITAL' || instancia.atributo.substring(0, 2) === 'E_') ? 'LI_' : 'NI_') + (instancia.informacionSofrel < 10 ? '000' + instancia.informacionSofrel : instancia.informacionSofrel < 100 ? '00' + instancia.informacionSofrel : instancia.informacionSofrel < 1000 ? '0' + instancia.informacionSofrel : instancia.informacionSofrel) + '.Value'
    });
  });

  await workbook.xlsx.writeBuffer({
    based64: true
  }).then((xls64) => {
    var a = document.createElement("a");
    var data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(data);
    a.href = url;
    const today = new Date();
    a.download = 'ESAMUR - LS EDAR ' + nombreEdar + ' ' + today.getDate() + (today.getMonth() < 9 ? ('0' + (today.getMonth() + 1)) : (today.getMonth() + 1)) + today.getFullYear() + '.xlsx';
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