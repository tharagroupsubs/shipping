import { useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  Calculator, 
  Download, 
  CheckCircle2, 
  AlertCircle,
  Truck,
  MapPin,
  Scale
} from 'lucide-react';
import './App.css';

const ZONE_MAP = {
  // ZONE A: WITHIN STATE (KERALA)
  "KERALA": "A", "KL": "A", "KERLA": "A",
  
  // ZONE B: TAMIL NADU, KARNATAKA, PUDUCHERRY
  "TAMIL NADU": "B", "TN": "B", "TAMILNADU": "B",
  "KARNATAKA": "B", "KA": "B", "KARNATKA": "B",
  "PUDUCHERRY": "B", "PY": "B", "PONDICHERRY": "B",
  
  // ZONE C: AP, TELANGANA, GOA, MUMBAI, PUNE
  "ANDHRA PRADESH": "C", "AP": "C", "ANDHRA": "C",
  "TELANGANA": "C", "TS": "C", "TG": "C",
  "GOA": "C", "GA": "C",
  "MUMBAI": "C", "PUNE": "C", "BOMBAY": "C",
  
  // ZONE D: REST OF INDIA
  "DELHI": "D", "DL": "D", "NCR": "D",
  "UTTAR PRADESH": "D", "UP": "D",
  "RAJASTHAN": "D", "RJ": "D",
  "GUJARAT": "D", "GJ": "D",
  "MADHYA PRADESH": "D", "MP": "D",
  "JHARKHAND": "D", "JH": "D",
  "HARYANA": "D", "HR": "D",
  "WEST BENGAL": "D", "WB": "D",
  "MAHARASHTRA": "D", "MH": "D",
  "PUNJAB": "D", "PB": "D",
  "UTTARAKHAND": "D", "UK": "D", "UA": "D",
  "CHANDIGARH": "D", "CH": "D",
  "DADRA": "D", "DN": "D", "DAMAN": "D", "DIU": "D",

  // ZONE E: NORTH EAST & SPECIAL
  "ODISHA": "E", "OR": "E", "OD": "E", "ORISSA": "E",
  "CHATTISGARH": "E", "CG": "E", "CHHATTISGARH": "E",
  "MIZORAM": "E", "MZ": "E",
  "ASSAM": "E", "AS": "E",
  "NAGALAND": "E", "NL": "E",
  "BIHAR": "E", "BR": "E",
  "TRIPURA": "E", "TR": "E",
  "SIKKIM": "E", "SK": "E",
  "HIMACHAL PRADESH": "E", "HP": "E", "HIMACHAL": "E",
  "MANIPUR": "E", "MN": "E",

  // ZONE F: JAMMU & KASHMIR, MEGHALAYA, ARUNACHAL, LADAKH, ANDAMAN
  "JAMMU & KASHMIR": "F", "JK": "F", "J&K": "F", "JAMMU": "F", "KASHMIR": "F",
  "MEGHALAYA": "F", "ML": "F",
  "ARUNACHAL PRADESH": "F", "AR": "F", "ARUNACHAL": "F",
  "LADAKH": "F", "LA": "F",
  "ANDAMAN": "F", "NICOBAR": "F", "AN": "F",
  "LAKSHADWEEP": "F", "LD": "F",
};

// Default fixed rate sheet from business tariff (SURFACE)
const DEFAULT_SHIPPING_RATES = [
  { mode: 'SURFACE', slabUpper: 500, zoneCharges: { A: 55, B: 75, C: 80, D: 85, E: 95, F: 100 } },
  { mode: 'SURFACE', slabUpper: 1000, zoneCharges: { A: 75, B: 90, C: 95, D: 100, E: 145, F: 155 } },
  { mode: 'SURFACE', slabUpper: 1500, zoneCharges: { A: 110, B: 135, C: 140, D: 145, E: 200, F: 210 } },
  { mode: 'SURFACE', slabUpper: 2000, zoneCharges: { A: 140, B: 165, C: 180, D: 195, E: 240, F: 260 } },
  { mode: 'SURFACE', slabUpper: 2500, zoneCharges: { A: 170, B: 195, C: 220, D: 240, E: 280, F: 310 } },
  { mode: 'SURFACE', slabUpper: 3000, zoneCharges: { A: 200, B: 225, C: 260, D: 285, E: 320, F: 360 } },
  { mode: 'SURFACE', slabUpper: 3500, zoneCharges: { A: 230, B: 255, C: 300, D: 330, E: 360, F: 410 } },
  { mode: 'SURFACE', slabUpper: 4000, zoneCharges: { A: 260, B: 285, C: 320, D: 375, E: 400, F: 460 } },
  { mode: 'SURFACE', slabUpper: 4500, zoneCharges: { A: 290, B: 320, C: 360, D: 420, E: 440, F: 510 } },
  { mode: 'SURFACE', slabUpper: 5000, zoneCharges: { A: 320, B: 350, C: 400, D: 460, E: 480, F: 560 } },
];

const ALLOWED_STATUSES = new Set([
  'DELIVERED',
  'OUT_FOR_DELIVERY',
  'SHIPPED',
  'RETURNED_TO_ORIGIN',
  'RETURNING_TO_ORIGIN',
]);

const DOUBLE_CHARGE_STATUSES = new Set([
  'RETURNED_TO_ORIGIN',
  'RETURNING_TO_ORIGIN',
]);

const HEADER_CANDIDATES = {
  waybill: ['waybill', 'wbn', 'awb', 'tracking', 'trackingid', 'consignment', 'lrn', 'shipmentid'],
  mode: ['mode', 'servicetype', 'service', 'shippingmode', 'shipmentmode'],
  zone: ['zone', 'destinationzone', 'region'],
  state: ['state', 'destinationstate', 'to_state', 'to state'],
  paymentType: ['type', 'paymentmode', 'payment mode', 'cod/prepaid', 'shipmenttype', 'order type', 'category'],
  codAmount: ['codamount', 'cod amount', 'collectableamount', 'collectable amount', 'amount', 'orderamount', 'order amount', 'order value', 'value'],
  deadWeight: ['weight', 'deadweight', 'actualweight', 'shipmentweight', 'dead wt', 'dead_wt'],
  status: ['currentstatus', 'status', 'shipmentstatus'],
  internalWeight: ['internalweight', 'franchiseweight', 'internal wt', 'weight_internal_weight', 'internalwt', 'franchise wt', 'franchisewt', 'billedweight', 'billed weight', 'finalweight', 'final weight', 'revisedweight', 'revised weight'],
  c2cException: ['c2cweightexception', 'c2cexception', 'c2cweight', 'weight exception', 'c2c weight', 'exceptionweight', 'exception weight'],
  slab: ['weightslab', 'slab', 'weight', 'upto', 'range'],
};

const normalizeKey = (value) => String(value ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');

// Normalize waybill: stringify, strip trailing .0 from Excel numeric reads, remove spaces
const normalizeWaybill = (value) => {
  if (value === null || value === undefined) return '';
  return String(value).trim().replace(/\.0+$/, '').replace(/\s+/g, '');
};

const toUpperClean = (value) => String(value ?? '').trim().toUpperCase();

const parseNumeric = (value) => {
  if (value === null || value === undefined || value === '') return null;
  const clean = String(value).replace(/,/g, '').match(/[0-9]*\.?[0-9]+/g);
  if (!clean) return null;
  const parsed = Number(clean[0]);
  if (Number.isNaN(parsed)) return null;
  return parsed;
};

const normalizeWeightToGram = (value) => {
  const num = parseNumeric(value);
  if (num === null) return null;
  if (num > 0 && num < 50) {
    return Math.round(num * 1000);
  }
  return Math.round(num);
};

const findKeyInRow = (row, candidates) => {
  const keys = Object.keys(row || {});
  const normalizedCandidates = candidates.map((item) => normalizeKey(item));

  for (const key of keys) {
    const normalized = normalizeKey(key);
    if (normalizedCandidates.includes(normalized)) return key;
  }

  for (const key of keys) {
    const normalized = normalizeKey(key);
    if (normalizedCandidates.some((candidate) => normalized.includes(candidate) || candidate.includes(normalized))) {
      return key;
    }
  }

  return null;
};

const getRowValue = (row, candidates) => {
  const key = findKeyInRow(row, candidates);
  return key ? row[key] : null;
};

const extractZone = (row) => {
  const zoneValue = toUpperClean(getRowValue(row, HEADER_CANDIDATES.zone));
  if (zoneValue) {
    const normalizedZone = zoneValue.replace(/^ZONE\s*/i, '').trim();
    if (/^[A-F]$/.test(normalizedZone)) return normalizedZone;
    if (ZONE_MAP[normalizedZone]) return ZONE_MAP[normalizedZone];
  }

  const stateValue = toUpperClean(getRowValue(row, HEADER_CANDIDATES.state));
  if (!stateValue) return '';
  const direct = ZONE_MAP[stateValue];
  if (direct) return direct;
  const fuzzy = Object.keys(ZONE_MAP).find((state) => stateValue.includes(state));
  return fuzzy ? ZONE_MAP[fuzzy] : '';
};

const getMatchedSlab = (sortedSlabs, weightInGram) => {
  if (!sortedSlabs.length) return null;
  if (weightInGram === null) return sortedSlabs[0];
  return sortedSlabs.find((item) => weightInGram <= item.slabUpper) || sortedSlabs[sortedSlabs.length - 1];
};

function App() {
  const [shipmentRows, setShipmentRows] = useState([]);
  const [weightRows, setWeightRows] = useState([]);
  const [fileNames, setFileNames] = useState({
    shipments: '',
    weights: '',
  });
  const [isProcessing, setIsProcessing] = useState(false);
  const [status, setStatus] = useState('pending');
  const [summary, setSummary] = useState(null);
  const [previewData, setPreviewData] = useState([]);
  const [errorMessage, setErrorMessage] = useState('');

  const shipmentInputRef = useRef(null);
  const weightInputRef = useRef(null);

  const resetApp = () => {
    if (shipmentInputRef.current) shipmentInputRef.current.value = '';
    if (weightInputRef.current) weightInputRef.current.value = '';

    setShipmentRows([]);
    setWeightRows([]);
    setFileNames({ shipments: '', weights: '' });
    setStatus('pending');
    setSummary(null);
    setPreviewData([]);
    setErrorMessage('');
  };

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    setErrorMessage('');
    setStatus('pending');
    setSummary(null);
    setPreviewData([]);
    setFileNames((prev) => ({ ...prev, [type]: file.name }));

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (data.length === 0) {
          alert("The uploaded file seems to be empty.");
          return;
        }

        if (type === 'shipments') setShipmentRows(data);
        if (type === 'weights') setWeightRows(data);
      } catch (err) {
        alert("Error reading excel file. Please ensure it's a valid .xlsx or .csv file.");
        console.error(err);
      }
    };
    reader.readAsBinaryString(file);
  };

  const canProcess = shipmentRows.length > 0 && weightRows.length > 0;

  const calculateShipping = () => {
    if (!canProcess) {
      alert('Please upload Shipment Sheet and Weight Comparison Sheet.');
      return;
    }

    setIsProcessing(true);
    setErrorMessage('');
    
    try {
      const weightMap = new Map();
      weightRows.forEach((row) => {
        const waybillRaw = getRowValue(row, HEADER_CANDIDATES.waybill);
        const waybill = normalizeWaybill(waybillRaw);
        if (!waybill) return;
        weightMap.set(waybill, row);
      });

      // Debug: log first weight row keys to console for diagnosis
      if (weightRows.length > 0) {
        console.log('[WeightSheet] First row keys:', Object.keys(weightRows[0]));
        console.log('[WeightSheet] Sample WBN:', normalizeWaybill(getRowValue(weightRows[0], HEADER_CANDIDATES.waybill)));
        console.log('[WeightSheet] Sample internalWeight:', getRowValue(weightRows[0], HEADER_CANDIDATES.internalWeight));
      }

      let matchedWeightRowsCount = 0;
      let eligibleStatusesCount = 0;

      const output = shipmentRows.map((shipment) => {
        const waybill = normalizeWaybill(getRowValue(shipment, HEADER_CANDIDATES.waybill));
        const mode = toUpperClean(getRowValue(shipment, HEADER_CANDIDATES.mode)) || 'SURFACE';
        const zone = extractZone(shipment);
        const statusValue = toUpperClean(getRowValue(shipment, HEADER_CANDIDATES.status)).replace(/[\s-]+/g, '_');
        const paymentType = toUpperClean(getRowValue(shipment, HEADER_CANDIDATES.paymentType));
        const codAmount = parseNumeric(getRowValue(shipment, HEADER_CANDIDATES.codAmount)) ?? 0;
        const deadWeight = normalizeWeightToGram(getRowValue(shipment, HEADER_CANDIDATES.deadWeight));

        const matchedWeightRow = waybill ? weightMap.get(normalizeWaybill(waybill)) : null;
        const internalWeight = normalizeWeightToGram(getRowValue(matchedWeightRow || {}, HEADER_CANDIDATES.internalWeight));
        const c2cWeightException = normalizeWeightToGram(getRowValue(matchedWeightRow || {}, HEADER_CANDIDATES.c2cException));

        const effectiveWeight = c2cWeightException ?? internalWeight ?? deadWeight;

        const modeRates = DEFAULT_SHIPPING_RATES.filter((rate) => rate.mode === mode);
        const ratesToUse = modeRates.length ? modeRates : DEFAULT_SHIPPING_RATES;
        const slab = getMatchedSlab(ratesToUse, effectiveWeight);

        let normalRate = null;
        if (slab && zone && slab.zoneCharges[zone] !== undefined) {
          normalRate = slab.zoneCharges[zone];
        }

        let shippingCharge = '';
        const isValidStatus = ALLOWED_STATUSES.has(statusValue);
        const isCodShipment = paymentType.includes('COD');
        const isRtoStatus = DOUBLE_CHARGE_STATUSES.has(statusValue);
        if (isValidStatus && typeof normalRate === 'number') {
          const fixedShippingRate = isRtoStatus ? normalRate * 2 : normalRate;

          if (isCodShipment) {
            // COD fee is not applied for RTO statuses.
            if (isRtoStatus) {
              shippingCharge = fixedShippingRate;
            } else {
              const codFee = codAmount > 2000 ? codAmount * 0.02 : 40;
              shippingCharge = fixedShippingRate + codFee;
            }
          } else {
            shippingCharge = fixedShippingRate;
          }
        }

        const chargedWeight = slab ? slab.slabUpper : '';

        if (matchedWeightRow) matchedWeightRowsCount += 1;
        if (isValidStatus) eligibleStatusesCount += 1;

        // Keep all uploaded columns and only update/add the requested output columns.
        return {
          ...shipment,
          'Current Status': statusValue,
          'Zone': zone || 'UNKNOWN',
          'Charged Weight': chargedWeight,
          'Shipping Charges': shippingCharge,
        };
      });

      const totalShipping = output.reduce((sum, item) => {
        const value = parseNumeric(item['Shipping Charges']);
        return sum + (value ?? 0);
      }, 0);

      const chargeableRows = output.filter((row) => typeof row['Shipping Charges'] === 'number');

      setSummary({
        totalRows: output.length,
        matchedWeightRows: matchedWeightRowsCount,
        eligibleStatuses: eligibleStatusesCount,
        billedRows: chargeableRows.length,
        totalShipping,
      });

      setPreviewData(output.slice(0, 8));
      downloadExcel(output);
      setStatus('success');
    } catch (error) {
      console.error(error);
      setErrorMessage(`An error occurred while processing files: ${error.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadExcel = (data) => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Shipping Costs");
    
    const date = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `shipping_costs_${date}.xlsx`);
  };

  return (
    <div className="container">
      <header className="header">
        <div className={`status-badge ${canProcess ? 'ready' : ''}`}>
          <Truck size={14} style={{marginRight: '6px'}} /> Automated Logistics Engine
        </div>
        <h1>Shipping Calculator</h1>
        <p>Fixed rate + shipment + weight comparison analyzer</p>
      </header>

      <main className="glass-card">
        <div className="upload-section">
          <div className={`upload-box wide ${shipmentRows.length > 0 ? 'active' : ''}`}>
            <input
              ref={shipmentInputRef}
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={(event) => handleFileUpload(event, 'shipments')}
            />
            <Upload className="upload-icon" />
            <h3>Shipment Input Sheet</h3>
            <p>Waybill/WBN, mode, zone/state, dead weight, current status</p>
            {fileNames.shipments && (
              <div className="file-info">
                <CheckCircle2 size={16} /> {fileNames.shipments}
              </div>
            )}
          </div>
        </div>

        <div className="upload-section">
          <div className={`upload-box wide ${weightRows.length > 0 ? 'active' : ''}`}>
            <input
              ref={weightInputRef}
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={(event) => handleFileUpload(event, 'weights')}
            />
            <Upload className="upload-icon" />
            <h3>Weight Comparison Sheet</h3>
            <p>Waybill/WBN, internal weight, C2C weight exception</p>
            {fileNames.weights && (
              <div className="file-info">
                <CheckCircle2 size={16} /> {fileNames.weights}
              </div>
            )}
          </div>
        </div>

        <div className="rates-preview">
          <h4>Default Fixed Rate Sheet (SURFACE)</h4>
          <div className="rates-table-container">
            <table className="rates-table">
              <thead>
                <tr>
                  <th>Weight</th>
                  <th>Zone A</th>
                  <th>Zone B</th>
                  <th>Zone C</th>
                  <th>Zone D</th>
                  <th>Zone E</th>
                  <th>Zone F</th>
                </tr>
              </thead>
              <tbody>
                {DEFAULT_SHIPPING_RATES.map((rate, idx) => (
                  <tr key={idx}>
                    <td>{rate.slabUpper}</td>
                    <td>{rate.zoneCharges.A}</td>
                    <td>{rate.zoneCharges.B}</td>
                    <td>{rate.zoneCharges.C}</td>
                    <td>{rate.zoneCharges.D}</td>
                    <td>{rate.zoneCharges.E}</td>
                    <td>{rate.zoneCharges.F}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {errorMessage && (
          <div className="debug-log" style={{ marginTop: 0, marginBottom: '1.5rem' }}>
            <div className="log-entry error" style={{ marginBottom: 0 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <AlertCircle size={14} color="var(--error)" />
                <span style={{ color: 'var(--error)' }}>{errorMessage}</span>
              </div>
            </div>
          </div>
        )}

        {summary && status === 'success' && (
          <div className="summary-cards">
            <div className="summary-card">
              <span className="summary-label">Total Shipments</span>
              <span className="summary-value">{summary.totalRows}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">Matched Weights</span>
              <span className="summary-value">{summary.matchedWeightRows}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">Eligible Status Rows</span>
              <span className="summary-value">{summary.eligibleStatuses}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">Rows With Charges</span>
              <span className="summary-value">{summary.billedRows}</span>
            </div>
            <div className="summary-card">
              <span className="summary-label">Total Shipping</span>
              <span className="summary-value">₹{summary.totalShipping.toLocaleString()}</span>
            </div>
          </div>
        )}

        {previewData.length > 0 && (
          <div className="results-preview">
            <h4>Processing Preview (First 8 Rows)</h4>
            <div className="rates-table-container">
              <table className="rates-table">
                <thead>
                  <tr>
                    <th>Waybill / WBN</th>
                    <th>Current Status</th>
                    <th>Zone</th>
                    <th>Charged Weight</th>
                    <th>Shipping Charges</th>
                  </tr>
                </thead>
                <tbody>
                  {previewData.map((row, idx) => (
                    <tr key={idx}>
                      <td>{getRowValue(row, HEADER_CANDIDATES.waybill) || row['Waybill / WBN'] || '-'}</td>
                      <td>{row['Current Status'] || '-'}</td>
                      <td>{row['Zone'] || '-'}</td>
                      <td>{row['Charged Weight'] || '-'}</td>
                      <td style={{ fontWeight: '700', color: typeof row['Shipping Charges'] === 'number' ? 'var(--success)' : 'var(--text-muted)' }}>
                        {typeof row['Shipping Charges'] === 'number' ? `₹${row['Shipping Charges']}` : '-'}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        <div className="button-group">
          <button 
            className={`btn-primary ${status === 'success' ? 'btn-success' : ''}`} 
            onClick={calculateShipping}
            disabled={!canProcess || isProcessing}
          >
            {isProcessing ? (
              <><div className="spinner"></div> Processing...</>
            ) : (
              <>
                {status === 'success' ? <Download size={20} /> : <Calculator size={20} />}
                {status === 'success' ? 'Recalculate & Download Again' : 'Compare & Download Excel'}
              </>
            )}
          </button>
          
          {(canProcess || status === 'success') && (
            <button className="btn-secondary" onClick={resetApp}>
              <Upload size={18} /> Upload New Files
            </button>
          )}
        </div>

        {!canProcess && (
          <div className="debug-log" style={{ marginTop: '1.5rem' }}>
            <div className="log-entry" style={{ marginBottom: 0 }}>
              Upload Shipment Sheet and Weight Comparison Sheet to enable processing.
            </div>
          </div>
        )}
      </main>

      <footer style={{marginTop: '3rem', color: 'var(--text-muted)', fontSize: '0.85rem', textAlign: 'center'}}>
        <p>© 2026 Antigravity Logistics Solutions. All rights reserved.</p>
        <div style={{display: 'flex', justifyContent: 'center', gap: '1.5rem', marginTop: '1rem'}}>
          <span style={{display: 'flex', alignItems: 'center', gap: '4px'}}><Scale size={14}/> Accurate Weight Slabs</span>
          <span style={{display: 'flex', alignItems: 'center', gap: '4px'}}><MapPin size={14}/> Multi-Zone Routing</span>
        </div>
      </footer>
    </div>
  );
}

export default App;

