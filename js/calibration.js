let mean = [];
let sigma = [];
let nominal = [];
let alpha = 0;
let alphaSigma = 0;
let validDimensions = false;
let validSkew = false;
let expectedShrinkage = 0;
let xerror = 0;
let xsigma = 0;
let yerror = 0;
let ysigma = 0;

const range = (a, b) => Array(b - a + 1).fill(a).map((x, y) => x + y);
const dot = (a, b) => a.map((x, i) => a[i] * b[i]).reduce((m, n) => m + n, 0);
const mult = (a, b) => a.map((x, i) => x * b[i]);

const calculateSum = (arr) => {
    return arr.reduce((total, current) => isNaN(current) ? total : total + current, 0);
};

let dataSchema = [
    ['Outer', 'X', '1', '', '', '', '120', ''],
    ['Outer', 'X', '2', '', '', '', '100', ''],
    ['Outer', 'X', '3', '', '', '', '80' , ''],
    ['Outer', 'X', '4', '', '', '', '60' , ''],
    ['Outer', 'X', '5', '', '', '', '40' , ''],
    ['Inner', 'X', '1', '', '', '', '120', ''],
    ['Inner', 'X', '2', '', '', '', '100', ''],
    ['Inner', 'X', '3', '', '', '', '80' , ''],
    ['Inner', 'X', '4', '', '', '', '60' , ''],
    ['Inner', 'X', '5', '', '', '', '40' , ''],
    ['Outer', 'Y', '1', '', '', '', '120', ''],
    ['Outer', 'Y', '2', '', '', '', '100', ''],
    ['Outer', 'Y', '3', '', '', '', '80' , ''],
    ['Outer', 'Y', '4', '', '', '', '60' , ''],
    ['Outer', 'Y', '5', '', '', '', '40' , ''],
    ['Inner', 'Y', '1', '', '', '', '120', ''],
    ['Inner', 'Y', '2', '', '', '', '100', ''],
    ['Inner', 'Y', '3', '', '', '', '80' , ''],
    ['Inner', 'Y', '4', '', '', '', '60' , ''],
    ['Inner', 'Y', '5', '', '', '', '40' , ''],
    ['Outer', 'D' ,'1', '', '', '', '120', ''],
    ['Outer', 'D' ,'2', '', '', '', '100', ''],
    ['Outer', 'D' ,'3', '', '', '', '80' , ''],
    ['Outer', 'D' ,'4', '', '', '', '60' , ''],
    ['Outer', 'd' ,'1', '', '', '', '120', ''],
    ['Outer', 'd' ,'2', '', '', '', '100', ''],
    ['Outer', 'd' ,'3', '', '', '', '80' , ''],
    ['Outer', 'd' ,'4', '', '', '', '60' , '']
];

let worksheet = jspreadsheet(document.getElementById('spreadsheet'), {
    data: dataSchema,
    columns: [
        { type: 'text', title:'Type', width:80, readOnly: true },
        { type: 'text', title:'Axis', width:60, readOnly: true  },
        { type: 'text', title:'Location',  width:80, readOnly: true },
        { type: 'numeric', title:'Meas 1', width:90 },
        { type: 'numeric', title:'Meas 2', width:90 },
        { type: 'numeric', title:'Meas 3', width:90 },
        { type: 'numeric', title:'Nominal (mm)', width:120, readOnly: true },
        { type: 'numeric', title:'Avg. (mm)', width:100 , readOnly: true, format: '#,##0.00' }
     ],
     nestedHeaders:[
        [
            { title: 'Configuration Details', colspan: '3' },
            { title: 'Physical Metrology Checks', colspan: '5' }
        ],
     ],
     tableOverflow: true,
     tableHeight: '380px'
});
worksheet.hideIndex();

function hideMeasurementRows() {
    let M = Number(document.getElementById('numMeasPts').value);
    for (let ii = 0; ii < 28; ii++) worksheet.showRow(ii);
    if (M < 5) [4, 9, 14, 19, 23, 27].forEach(r => worksheet.hideRow(r));
    if (M < 4) [3, 8, 13, 18, 22, 26].forEach(r => worksheet.hideRow(r));
    if (M < 3) [2, 7, 12, 17, 21, 25].forEach(r => worksheet.hideRow(r));
}

function extractColumnVector(rows, col) {
    return rows.map(r => Number(worksheet.getValueFromCoords(col, r)));
}

function extractRowVector(row, cols) {
    return cols.map(c => Number(worksheet.getValueFromCoords(c, row)));
}

function setNominals() {
    let M = Number(document.getElementById('numMeasPts').value);
    let N = Number(document.getElementById('calistarSize').value);
    let printScale = Number(document.getElementById('printScale').value);
    let nomCol = 6;

    for (let ii = 0; ii < 5; ii++) {
        let val = Math.max((N - N * ii / M) * printScale, 0);
        let strVal = val === 0 ? NaN : String(val);
        [0, 5, 10, 15].forEach(offset => worksheet.setValueFromCoords(nomCol, ii + offset, strVal, true));
    }

    for (let ii = 0; ii < 4; ii++) {
        let val = Math.max((N - N * ii / M) * printScale, 0);
        let strVal = (val === 0 || ii >= M - 1) ? 0 : String(val);
        worksheet.setValueFromCoords(nomCol, ii + 20, strVal, true);
        worksheet.setValueFromCoords(nomCol, ii + 24, strVal, true);
    }
    nominal = extractColumnVector(range(0, 27), nomCol);
}

function handleParamChange() {
    hideMeasurementRows();
    setNominals();
}

let numValidMeasurements = [];

function updateMean() {
    numValidMeasurements = range(0, 27).map(ii => {
        let w = extractRowVector(ii, [3, 4, 5]);
        return w.filter(x => !isNaN(x) && x > 0).length;
    });

    mean = range(0, 27).map((ii) => {
        let w = extractRowVector(ii, [3, 4, 5]);
        let validVals = w.filter(x => !isNaN(x) && x > 0);
        let m = validVals.length > 0 ? calculateSum(validVals) / validVals.length : 0;
        if (m > 0) {
            worksheet.setValueFromCoords(7, ii, m.toFixed(2), true);
        } else {
            worksheet.setValueFromCoords(7, ii, '', true);
        }
        return m;
    });
}

function updateSigma() {
    updateMean();
    sigma = range(0, 27).map((ii) => {
        let w = extractRowVector(ii, [3, 4, 5]).filter(x => !isNaN(x) && x > 0);
        if (w.length <= 1) return 0;
        let diffs = w.map(x => (x - mean[ii]) * (x - mean[ii]));
        return Math.sqrt(calculateSum(diffs) / (w.length - 1));
    });
}

function calculateDimensionality(rows) {
    let numValid = 0, sum = 0, sumsq = 0;
    let sigma0 = document.getElementById('sample_error').checked ? Number(document.getElementById('caliper_error').value) : 0;

    rows.forEach(row => {
        let nom = nominal[row];
        if (mean[row] > 0 && nom > 0) {
            sum += (mean[row] - nom) / nom;
            let sigma2 = sigma[row] * sigma[row] + sigma0 * sigma0;
            numValid++;
            sumsq += sigma2 / (nom * nom);
        }
    });
    return numValid > 0 ? [sum / numValid, Math.sqrt(sumsq / (numValid * numValid))] : [0, 0];
}

function updateDimensionality() {
    let xResults = calculateDimensionality(range(0, 9));
    xerror = xResults[0]; xsigma = xResults[1];
    document.getElementById('xerror').innerHTML = (xerror * 100).toFixed(3) + '%';
    document.getElementById('xcorrection').innerHTML = (-xerror * 100).toFixed(3) + '%';
    document.getElementById('xNumSigma').innerHTML = xsigma > 0 ? (xerror / xsigma).toFixed(1) : '0.0';

    let yResults = calculateDimensionality(range(10, 19));
    yerror = yResults[0]; ysigma = yResults[1];
    document.getElementById('yerror').innerHTML = (yerror * 100).toFixed(3) + '%';
    document.getElementById('ycorrection').innerHTML = (-yerror * 100).toFixed(3) + '%';
    document.getElementById('yNumSigma').innerHTML = ysigma > 0 ? (yerror / ysigma).toFixed(1) : '0.0';

    document.getElementById('yourCorrection').innerHTML = `(${(xerror * 100).toFixed(2)}%, ${(yerror * 100).toFixed(2)}%)`;

    document.getElementById('slicerscalingX').innerHTML = xerror !== 0 ? ((1 / (1 + xerror)) * 100).toFixed(3) : '100.000';
    document.getElementById('slicerscalingY').innerHTML = yerror !== 0 ? ((1 / (1 + yerror)) * 100).toFixed(3) : '100.000';

    let dGraph = document.getElementById('dimGraph');
    dGraph.data[0].x = [yerror * 100, xerror * 100];
    dGraph.data[0].error_x.array = [2 * ysigma * 100, 2 * xsigma * 100];
    dGraph.data[1].x = [yerror * 100, xerror * 100];
    dGraph.data[1].error_x.array = [ysigma * 100, xsigma * 100];
    Plotly.redraw('dimGraph');
}

function computeTotalLength(rows) {
    let sum = 0, nominalSum = 0;
    rows.forEach(row => {
        if (numValidMeasurements[row] > 0 && nominal[row] > 0) {
            sum += mean[row];
            nominalSum += nominal[row];
        }
    });
    return nominalSum > 0 ? sum / nominalSum : 0;
}

function calculateSkew() {
    let p = computeTotalLength([20, 21, 22, 23]);
    let q = computeTotalLength([24, 25, 26, 27]);
    let a = computeTotalLength([0, 1, 2, 3, 4]) * Math.sqrt(2) / 2;
    let b = computeTotalLength([10, 11, 12, 13, 14]) * Math.sqrt(2) / 2;

    if (p === 0 || q === 0 || a === 0 || b === 0) return [NaN, 0];

    let cosAlpha = (p * p - q * q) / (4 * a * b);
    cosAlpha = Math.max(-1, Math.min(1, cosAlpha)); 
    let alphaRad = Math.acos(cosAlpha);
    alpha = alphaRad * 180 / Math.PI;

    let sinAlpha = Math.sin(alphaRad);
    let sigma0 = document.getElementById('sample_error').checked ? Number(document.getElementById('caliper_error').value) : 0;
    let sig2p = (sigma0 * sigma0) / (p*p);
    let sig2q = (sigma0 * sigma0) / (q*q);

    let gradp = Math.abs(p / (2 * a * b * (sinAlpha || 1)));
    let gradq = Math.abs(q / (2 * a * b * (sinAlpha || 1)));

    alphaSigma = Math.sqrt(gradp * gradp * sig2p + gradq * gradq * sig2q) * 180 / Math.PI;
    return [alpha, alphaSigma];
}

function updateSkew() {
    let v = calculateSkew();
    alpha = v[0]; alphaSigma = v[1];

    if (isNaN(alpha)) {
        document.getElementById('parAngle').innerHTML = 'N/A';
        document.getElementById('skewAngle').innerHTML = 'N/A';
        document.getElementById('skewCorrection').innerHTML = 'N/A';
        document.getElementById('skewNumSigmas').innerHTML = '0.00';
        return;
    }

    document.getElementById('parAngle').innerHTML = alpha.toFixed(4) + '°';
    document.getElementById('skewAngle').innerHTML = (alpha - 90).toFixed(4) + '°';
    document.getElementById('skewCorrection').innerHTML = (90 - alpha).toFixed(4) + '°';
    document.getElementById('skewNumSigmas').innerHTML = alphaSigma > 0 ? Math.abs((90 - alpha) / alphaSigma).toFixed(2) : '0.00';

    let sGraph = document.getElementById('skewGraph');
    sGraph.data[0].x = [alpha - 90];
    sGraph.data[0].error_x.array = [2 * alphaSigma];
    sGraph.data[1].x = [alpha - 90];
    sGraph.data[1].error_x.array = [alphaSigma];
    Plotly.redraw('skewGraph');
}

function validateMeasurements() {
    let validx = mean.slice(0, 10).some(x => x > 0);
    let validy = mean.slice(10, 20).some(x => x > 0);
    let validD = mean.slice(20, 24).some(x => x > 0);
    let validd = mean.slice(24, 28).some(x => x > 0);

    validDimensions = validx && validy;
    validSkew = validDimensions && validD && validd;

    document.getElementById('validDimensions').style.display = validDimensions ? "none" : "";
    document.getElementById('validSkew').style.display = validSkew ? "none" : "";
    
    if (validDimensions && validSkew) {
        document.getElementById('measurementWarnings').style.display = "none";
        document.getElementById('valideverything').style.display = "";
    } else {
        document.getElementById('measurementWarnings').style.display = "";
        document.getElementById('valideverything').style.display = "none";
    }
}

function updateMaterial() {
    let material = document.getElementById('materials').value;
    switch (material) {
        case 'PLA':  expectedShrinkage = 0.003; break;
        case 'PETG': expectedShrinkage = 0.008; break;
        case 'ABS':  expectedShrinkage = 0.016; break;
        default:     expectedShrinkage = 0.000;
    }
    document.getElementById('expectedShrinkage').innerHTML = (-expectedShrinkage * 100).toFixed(2) + '%';
}

function processMetrics() {
    updateSigma();
    validateMeasurements();
    updateDimensionality();
    updateSkew();
}

function initPlots() {
    let layoutConfig = {
        margin: { l: 60, r: 20, t: 30, b: 40 },
        xaxis: { showgrid: false, zeroline: true },
        yaxis: { showgrid: false, showline: false }
    };

    Plotly.newPlot('dimGraph', [
        { x: [0, 0], y: ['Y Axis', 'X Axis'], mode: 'markers', name: '2&sigma;', error_x: { type: 'data', array: [0, 0], color: '#333', width: 5 }, type: 'scatter', marker: { size: 1, color: '#333' } },
        { x: [0, 0], y: ['Y Axis', 'X Axis'], mode: 'markers', name: '1&sigma;', error_x: { type: 'data', array: [0, 0], color: '#3f51b5', thickness: 4 }, type: 'scatter', marker: { size: 10, color: '#3f51b5' } }
    ], { ...layoutConfig, title: 'Relative Dimensional Error (%)' }, { displayModeBar: false });

    Plotly.newPlot('skewGraph', [
        { x: [0], y: ['Skew'], mode: 'markers', name: '2&sigma;', error_x: { type: 'data', array: [0], color: '#333', width: 5 }, type: 'scatter', marker: { size: 1, color: '#333' } },
        { x: [0], y: ['Skew'], mode: 'markers', name: '1&sigma;', error_x: { type: 'data', array: [0], color: '#3f51b5', thickness: 4 }, type: 'scatter', marker: { size: 10, color: '#3f51b5' } }
    ], { ...layoutConfig, title: 'Skew Error Metric (Degrees)' }, { displayModeBar: false });
}

window.onload = () => {
    initPlots();
    handleParamChange();
    updateMaterial();
};
