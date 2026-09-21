import { useEffect, useState } from "react";

const API = "http://localhost:5000"; // backend ka base URL

export default function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [dataFile, setDataFile]         = useState(null);
  const [files, setFiles]               = useState([]);
  const [loading, setLoading]           = useState(false);
  const [msg, setMsg]                   = useState("");

  // page load pe purani files dikha do
  useEffect(() => {
    fetchFiles();
  }, []);

  const fetchFiles = async () => {
    try {
      const res  = await fetch(`${API}/api/files`);
      const json = await res.json();
      setFiles(json.files || []);
    } catch {
      /* ignore */
    }
  };

  const handleGenerate = async () => {
    if (!templateFile || !dataFile) {
      setMsg("⚠️ Dono Excel files select karo (Template + Data).");
      return;
    }

    const fd = new FormData();
    fd.append("template", templateFile);
    fd.append("data", dataFile);

    setLoading(true);
    setMsg("");
    try {
      const res  = await fetch(`${API}/api/generate`, {
        method: "POST",
        body: fd,
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || "Generation failed");

      setFiles(json.files || []);
      setMsg(`✅ ${json.count} file(s) generate ho gayi OutputFolder me.`);
    } catch (err) {
      setMsg("❌ Error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleReset = async () => {
    await fetch(`${API}/api/reset`, { method: "POST" });
    setFiles([]);
    setMsg("🧹 OutputFolder clear ho gaya.");
  };

  return (
    <div className="wrap">
      <h1>📦 Inv Slip Generator</h1>
      <p className="sub">
        Template Excel + Data Excel upload karo — har row ke liye alag
        Excel file ban jaayegi (naam = Inv No).
      </p>

      <div className="card">
        <label className="field">
          <span>1️⃣ Template File (Inv Smp.xlsx)</span>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setTemplateFile(e.target.files[0])}
          />
        </label>

        <label className="field">
          <span>2️⃣ Data File (Smaple_data.xlsx)</span>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(e) => setDataFile(e.target.files[0])}
          />
        </label>

        <div className="actions">
          <button onClick={handleGenerate} disabled={loading}>
            {loading ? "Generating..." : "🚀 Generate Files"}
          </button>
          <button className="ghost" onClick={handleReset} disabled={loading}>
            🧹 Reset
          </button>
          {files.length > 0 && (
            <a className="zip" href={`${API}/api/download-all`}>
              ⬇️ Download All (.zip)
            </a>
          )}
        </div>

        {msg && <div className="msg">{msg}</div>}
      </div>

      {files.length > 0 && (
        <div className="card">
          <h3>📁 Generated Files ({files.length})</h3>
          <ul className="filelist">
            {files.map((f) => (
              <li key={f}>
                <span>{f}</span>
                <a href={`${API}/api/download/${encodeURIComponent(f)}`}>
                  Download
                </a>
              </li>
            ))}
          </ul>
        </div>
      )}
    </div>
  );
}