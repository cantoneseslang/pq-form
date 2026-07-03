import {
  buildDailyReportRowValues,
  transferMinutes,
  formatDailyProductName,
  pickTargetRow,
  resolveMoldingMachineBlock,
} from '../lib/dailyReport.js';

const productTypes = {
  地槽: true,
  企筒: false,
  批灰角: false,
};

const ut075Main = [
  { operator: '達', start: '09:05', finish: '10:30', speed: '轉機', thickness: '0.8', width: '75', height: '50', length: '2440' },
  { operator: '嫻', load: '11:40', start: '11:50', finish: '12:25', speed: '80', thickness: '0.8', width: '75', height: '50', length: '2440' },
];
const ut075Mat = [{ qty: '200' }];

const ut100Main = [
  { operator: '達', start: '13:35', finish: '16:00', speed: '轉機', thickness: '0.8', width: '100', height: '50', length: '2440' },
];
const ut100Mat = [{ qty: '' }];

const ut075 = buildDailyReportRowValues(ut075Main, ut075Mat, productTypes);
const ut100 = buildDailyReportRowValues(ut100Main, ut100Mat, productTypes);

console.log('UT075', ut075);
console.log('UT100', ut100);
console.log('transfer 10:10-13:40', transferMinutes('10:10', '13:40'));
console.log('product B', formatDailyProductName(ut075Main[0], productTypes));

const expected = {
  UT075: { A: '達/嫻', I: 85, J: 10, K: 35, L: '80', F: 200, H: 1 },
  UT100: { A: '達', I: 145, F: '', H: '', J: '', K: '', L: '' },
};

function check(name, got, exp) {
  for (const [k, v] of Object.entries(exp)) {
    const ok = String(got[k]) === String(v);
    console.log(`${name} ${k}: ${got[k]} expected ${v} ${ok ? 'OK' : 'FAIL'}`);
  }
}

check('UT075', ut075, expected.UT075);
check('UT100', ut100, expected.UT100);
