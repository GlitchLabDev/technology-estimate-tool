import { useEffect, useMemo, useState } from "react";
import { HyperFormula, type CellValue, type DetailedCellError } from "hyperformula";

interface FieldDef {
  row: number;
  label: string;
}

interface InputGroup {
  title: string;
  fields: FieldDef[];
}

const inputGroups: InputGroup[] = [
  {
    title: "Space & Seats",
    fields: [
      { row: 4, label: "USF (sq ft)" },
      { row: 5, label: "Seats" },
    ],
  },
  {
    title: "Office Size Breakdown",
    fields: [
      { row: 12, label: "120 sq ft offices" },
      { row: 13, label: "240 sq ft offices" },
      { row: 14, label: "360 sq ft offices" },
    ],
  },
  {
    title: "Conference & AV Rooms",
    fields: [
      { row: 18, label: "Phone Room / Focus VC Room" },
      { row: 19, label: "3-4 PAX" },
      { row: 20, label: "5-6 PAX" },
      { row: 21, label: "7-9 PAX" },
      { row: 22, label: "10-22 PAX" },
      { row: 24, label: "Client Facing 18-22 PAX" },
      { row: 26, label: "Hospitality Pantry" },
      { row: 27, label: "Pantry" },
      { row: 34, label: "Event / Training Room / MP Room" },
      { row: 35, label: "Cafeteria Screen (Not VC)" },
      { row: 36, label: "Lobby Screen" },
      { row: 37, label: "TV Distro" },
      { row: 38, label: "Office Screen (Seats / 20)" },
      { row: 39, label: "Broadcast Room" },
    ],
  },
  {
    title: "Support Rooms",
    fields: [
      { row: 40, label: "Tech Build Room" },
      { row: 41, label: "Tech HUB" },
      { row: 42, label: "IDF" },
      { row: 43, label: "Mail Room" },
    ],
  },
  {
    title: "Amenity Rooms",
    fields: [
      { row: 28, label: "Copy Room" },
      { row: 29, label: "Storage" },
      { row: 30, label: "Calm Room" },
      { row: 31, label: "Mothers Room" },
      { row: 32, label: "Multi Faith Room" },
    ],
  },
  {
    title: "Security",
    fields: [
      { row: 45, label: "Door" },
      { row: 46, label: "CCTV" },
      { row: 47, label: "Intercom" },
      { row: 48, label: "Security / AV Lic" },
      { row: 49, label: "Alert Enterprise VMS" },
      { row: 50, label: "Emergency Pull Cord (ADA Restroom)" },
      { row: 51, label: "Turnstile" },
      { row: 52, label: "Drive Gate" },
      { row: 53, label: "External Camera" },
      { row: 54, label: "Shooter Detection" },
    ],
  },
];

const allRows = inputGroups.flatMap((g) => g.fields.map((f) => f.row));
const defaultValues = allRows.reduce<Record<number, number>>((acc, r) => {
  acc[r] = 0;
  return acc;
}, {});

const isError = (v: CellValue): v is DetailedCellError =>
  v !== null && typeof v === "object" && "value" in v && typeof v.value === "string";

const safe = (v: CellValue): string | number | boolean | null => {
  if (isError(v)) return v.value;
  if (v === null || v === undefined || v === "") return 0;
  return v;
};

const safeNum = (v: CellValue): number => {
  const n = Number(safe(v));
  return isNaN(n) ? 0 : n;
};

const fmtCurrency = (v: CellValue) =>
  new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 2,
  }).format(safeNum(v));

const fmtNumber = (v: CellValue) =>
  new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(safeNum(v));

const SKIP_DESCRIPTIONS = [
  "Description",
  "Sub Total",
  "Total",
  "Non CAP Total Investment:",
  "Decom of Offices",
  "Shipping and Storage",
  "ManPower and T/E",
  "Workstream Total",
];

function App() {
  const [values, setValues] = useState<Record<number, number>>(defaultValues);
  const [hf, setHf] = useState<HyperFormula | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetch("/workbook.json")
      .then((res) => res.json())
      .then((data) => {
        const instance = HyperFormula.buildFromSheets(data, { licenseKey: "gpl-v3" });
        setHf(instance);
      })
      .catch((err) => setError(err.message));
  }, []);

  const summary = useMemo(() => {
    if (!hf) return null;
    const sheetId = hf.getSheetId("Tech Estimate Summary");
    if (sheetId === undefined) return null;
    const vals = hf.getSheetValues(sheetId);
    const categories = [
      { label: "Deskside - End User Compute", r: 7 },
      { label: "AV", r: 8 },
      { label: "Network", r: 9 },
      { label: "Server", r: 10 },
      { label: "Security", r: 11 },
      { label: "Circuits", r: 12 },
      { label: "Telecom", r: 13 },
      { label: "Tech Project Management", r: 14 },
    ];
    return {
      rows: categories.map((c) => ({
        label: c.label,
        budget: safeNum(vals[c.r - 1][4]),
        sqft: safeNum(vals[c.r - 1][6]),
        head: safeNum(vals[c.r - 1][8]),
      })),
      capTotal: safeNum(vals[14][4]),
      nonCapTotal: safeNum(vals[18][4]),
      projectTotal: safeNum(vals[20][4]),
    };
  }, [hf, values]);

  const details = useMemo(() => {
    if (!hf) return [];
    const sheetId = hf.getSheetId("Estimate Sheet");
    if (sheetId === undefined) return [];
    const vals = hf.getSheetValues(sheetId);
    const rows: {
      code: string;
      type: string;
      desc: string;
      qty: number;
      unit: number;
      cash: number;
    }[] = [];
    for (const row of vals) {
      const code = safe(row[2]);
      const type = safe(row[3]);
      const desc = safe(row[4]);
      const unit = safeNum(row[5]);
      const qty = safeNum(row[6]);
      const cash = safeNum(row[7]);
      if (cash > 0 && desc && !SKIP_DESCRIPTIONS.includes(String(desc).trim())) {
        rows.push({
          code: code === 0 || code === false ? "" : String(code),
          type: type === 0 || type === false ? "" : String(type),
          desc: String(desc),
          qty,
          unit,
          cash,
        });
      }
    }
    return rows;
  }, [hf, values]);

  const updateValue = (row: number, val: string) => {
    const num = Number(val) || 0;
    setValues((prev) => ({ ...prev, [row]: num }));
    if (hf) {
      const sheetId = hf.getSheetId("Test Fit");
      if (sheetId !== undefined) {
        hf.setCellContents({ sheet: sheetId, row: row - 1, col: 4 }, num);
      }
    }
  };

  const reset = () => {
    setValues(defaultValues);
    if (hf) {
      const sheetId = hf.getSheetId("Test Fit");
      if (sheetId !== undefined) {
        for (const row of allRows) {
          hf.setCellContents({ sheet: sheetId, row: row - 1, col: 4 }, 0);
        }
      }
    }
  };

  const exportCSV = () => {
    let csv = "Cost Code,Category,Description,Quantity,Unit Cost,Cash Value\n";
    for (const row of details) {
      const desc = String(row.desc).replace(/"/g, '""');
      const wrappedDesc =
        desc.includes(",") || desc.includes("\n") || desc.includes('"') ? `"${desc}"` : desc;
      csv += `${row.code},${row.type},${wrappedDesc},${row.qty},${row.unit.toFixed(2)},${row.cash.toFixed(2)}\n`;
    }
    const blob = new Blob([csv], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "technology_estimate.csv";
    a.click();
  };

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6 text-red-600">
        Error loading estimate engine: {error}
      </div>
    );
  }

  if (!hf) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6 text-slate-500">
        Loading estimate engine…
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 p-4 md:p-6">
      <div className="max-w-7xl mx-auto">
        <header className="mb-6">
          <h1 className="text-3xl font-bold text-slate-900">Technology Estimate Tool</h1>
          <p className="text-slate-600 mt-1">
            Enter your project test-fit data to generate a technology estimate.
          </p>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <section className="lg:col-span-1">
            <div className="bg-white rounded-xl shadow p-4">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-lg font-semibold">Project Inputs</h2>
                <button onClick={reset} className="text-sm text-blue-600 hover:underline">
                  Reset
                </button>
              </div>
              <div className="space-y-5 max-h-[80vh] overflow-y-auto pr-1">
                {inputGroups.map((group) => (
                  <div key={group.title}>
                    <h3 className="text-sm font-semibold text-slate-500 uppercase tracking-wide mb-2">
                      {group.title}
                    </h3>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      {group.fields.map((f) => (
                        <div key={f.row}>
                          <label
                            htmlFor={`input-${f.row}`}
                            className="block text-sm font-medium text-slate-700 mb-1"
                          >
                            {f.label}
                          </label>
                          <input
                            id={`input-${f.row}`}
                            type="number"
                            min="0"
                            step="1"
                            value={values[f.row] ?? 0}
                            onChange={(e) => updateValue(f.row, e.target.value)}
                            className="w-full border border-slate-300 rounded-md px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </section>

          <section className="lg:col-span-2 space-y-6">
            <div className="bg-white rounded-xl shadow p-4">
              <h2 className="text-lg font-semibold mb-4">Estimate Summary</h2>
              {summary && (
                <>
                  <table className="w-full text-sm">
                    <thead className="bg-slate-100">
                      <tr>
                        <th className="p-2 text-left">Category</th>
                        <th className="p-2 text-right">Budget</th>
                        <th className="p-2 text-right">$/SqFt</th>
                        <th className="p-2 text-right">$/Head</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y">
                      {summary.rows.map((row) => (
                        <tr key={row.label}>
                          <td className="p-2 font-medium">{row.label}</td>
                          <td className="p-2 text-right currency">{fmtCurrency(row.budget)}</td>
                          <td className="p-2 text-right currency">{fmtCurrency(row.sqft)}</td>
                          <td className="p-2 text-right currency">{fmtCurrency(row.head)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mt-4 pt-4 border-t">
                    <div className="bg-blue-50 p-3 rounded-lg">
                      <div className="text-xs text-slate-500 uppercase">Cap Total</div>
                      <div className="text-xl font-bold text-slate-900">
                        {fmtCurrency(summary.capTotal)}
                      </div>
                    </div>
                    <div className="bg-slate-100 p-3 rounded-lg">
                      <div className="text-xs text-slate-500 uppercase">Non-Cap Total</div>
                      <div className="text-xl font-bold text-slate-900">
                        {fmtCurrency(summary.nonCapTotal)}
                      </div>
                    </div>
                    <div className="bg-green-50 p-3 rounded-lg">
                      <div className="text-xs text-slate-500 uppercase">Project Total</div>
                      <div className="text-xl font-bold text-slate-900">
                        {fmtCurrency(summary.projectTotal)}
                      </div>
                    </div>
                  </div>
                </>
              )}
            </div>

            <div className="bg-white rounded-xl shadow p-4">
              <div className="flex justify-between items-center mb-4">
                <h2 className="text-lg font-semibold">Detailed Line Items</h2>
                <button
                  onClick={exportCSV}
                  className="bg-blue-600 text-white text-sm px-3 py-1.5 rounded hover:bg-blue-700"
                >
                  Export CSV
                </button>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-slate-100 text-left">
                    <tr>
                      <th className="p-2 font-medium">Cost Code</th>
                      <th className="p-2 font-medium">Category</th>
                      <th className="p-2 font-medium">Description</th>
                      <th className="p-2 font-medium text-right">Qty</th>
                      <th className="p-2 font-medium text-right">Unit Cost</th>
                      <th className="p-2 font-medium text-right">Cash Value</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {details.length === 0 ? (
                      <tr>
                        <td colSpan={6} className="p-4 text-center text-slate-500">
                          Enter project data to generate line items.
                        </td>
                      </tr>
                    ) : (
                      details.map((row, idx) => (
                        <tr key={idx}>
                          <td className="p-2">{row.code}</td>
                          <td className="p-2 text-slate-500">{row.type}</td>
                          <td className="p-2">{row.desc}</td>
                          <td className="p-2 text-right">{fmtNumber(row.qty)}</td>
                          <td className="p-2 text-right currency">{fmtCurrency(row.unit)}</td>
                          <td className="p-2 text-right currency font-medium">
                            {fmtCurrency(row.cash)}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        </div>
      </div>
    </div>
  );
}

export default App;
