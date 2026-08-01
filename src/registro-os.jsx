import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import {
  Zap, Plus, Trash2, Download, Camera, CheckCircle2, XCircle,
  LayoutGrid, ClipboardList, AlertTriangle,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  PieChart, Pie, Cell, LineChart, Line, Legend,
} from "recharts";

const EMPTY = {
  data: "",
  encarregado: "",
  os: "",
  tipoServico: "",
  regional: "",
  conformidade: "Conforme",
  descNaoConformidade: "",
  processo: "",
  nomeEletricistaLider: "",
  nomeEletricista: "",
  quemInspecionou: "",
  matriculaLider: "",
  matriculaEletricista: "",
  registroFoto: "Enviado",
  observacao: "",
  regrasOuro: {
    desligamentoRede: false,
    seccionamento: false,
    bloqueio: false,
    atestar: false,
    protegerEquipamentosEnergizados: false,
    epi: false,
  },
  regrasArquivos: {
    desligamentoRede: null,
    seccionamento: null,
    bloqueio: null,
    atestar: null,
    protegerEquipamentosEnergizados: null,
    epi: null,
  },
};

const RULES_DE_OURO = {
  desligamentoRede: "Desligamento de rede",
  seccionamento: "Seccionamento",
  bloqueio: "Bloqueio de energia",
  atestar: "Atestar ausência de tensão",
  protegerEquipamentosEnergizados: "Proteger equipamentos energizados",
  epi: "Uso de EPI adequado",
};

const RULE_IMAGES = {
  desligamentoRede: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M30 40h60v10H30zM30 60h50v10H30zM30 80h30v10H30z' fill='%23E8930C'/><circle cx='90' cy='70' r='10' fill='%23E8930C'/></svg>`,
  seccionamento: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M35 30l25 30-25 30' stroke='%23E8930C' stroke-width='10' fill='none'/><path d='M85 30l-25 30 25 30' stroke='%23E8930C' stroke-width='10' fill='none'/></svg>`,
  bloqueio: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><rect x='33' y='55' width='54' height='30' rx='6' fill='%23E8930C'/><path d='M48 55v-15a12 12 0 1 1 24 0v15' stroke='%2314181C' stroke-width='10' fill='none'/><circle cx='89' cy='70' r='5' fill='%2314181C'/></svg>`,
  atestar: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M30 70l20 18 40-52' stroke='%23E8930C' stroke-width='10' fill='none'/></svg>`,
  protegerEquipamentosEnergizados: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><rect x='35' y='45' width='50' height='30' rx='6' fill='%23E8930C'/><path d='M50 55h20M50 65h20M50 75h12' stroke='%2314181C' stroke-width='6' fill='none'/><circle cx='60' cy='35' r='12' fill='%23E8930C'/></svg>`,
  epi: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M38 80c0-25 22-32 22-32s22 7 22 32v10H38z' fill='%23E8930C'/><path d='M52 84h16M52 92h16' stroke='%2314181C' stroke-width='6' fill='none'/></svg>`,
};

const FOTO_STYLES = {
  "Enviado": { bg: "#EAF4EE", fg: "#1F6B3A", dot: "#2F9E52" },
  "Enviado com Não conformidade": { bg: "#FBF0E4", fg: "#9A5B14", dot: "#E8930C" },
  "Não enviado": { bg: "#FBEAEA", fg: "#A22E2E", dot: "#D64545" },
};

const CHART_COLORS = {
  conforme: "#2F9E52",
  naoConforme: "#D64545",
  accent: "#E8930C",
  accent2: "#4D8FFF",
};

const getRuleImage = (key) => RULE_IMAGES[key] || "";

const API_BASE = typeof window !== 'undefined' && window.location.hostname === 'localhost' 
  ? 'http://localhost:3000'
  : '';

export default function App() { 
  const [view, setView] = useState("registro"); // "registro" | "dashboard" | "regras"
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState(EMPTY);
  const [editingId, setEditingId] = useState(null);
  const [filter, setFilter] = useState("");
  const [monthFilter, setMonthFilter] = useState("all");
  const [dashboardMonth, setDashboardMonth] = useState("all");

  useEffect(() => {
    const fetchRows = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/records`);
        if (res.ok) {
          const data = await res.json();
          setRows(Array.isArray(data) ? data : []);
          return;
        }
      } catch (e) {
        // fallback to localStorage
      }
      const saved = window.localStorage.getItem("indicaServicosRows");
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed)) setRows(parsed);
        } catch (error) {
          console.warn("Falha ao carregar histórico de registros:", error);
        }
      }
    };
    fetchRows();
  }, []);

  useEffect(() => {
    window.localStorage.setItem("indicaServicosRows", JSON.stringify(rows));
  }, [rows]);

  const update = (field, value) => setForm((f) => ({ ...f, [field]: value }));

  const requiredOk =
    form.data && form.encarregado && form.os &&
    (form.conformidade !== "Não conforme" || form.descNaoConformidade.trim().length > 0);

  const handleSubmit = async () => {
    if (!requiredOk) return;
    const payload = editingId ? { ...form, id: editingId } : { ...form, id: crypto.randomUUID() };
    try {
      const res = await fetch(
        editingId ? `${API_BASE}/api/records/${payload.id}` : `${API_BASE}/api/records`,
        {
          method: editingId ? 'PUT' : 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        }
      );
      if (!res.ok) throw new Error('Falha ao salvar registro');
      const saved = await res.json();
      if (editingId) {
        setRows((rs) => rs.map((r) => (r.id === editingId ? saved : r)));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, saved]);
      }
      setForm(EMPTY);
    } catch (err) {
      console.error(err);
      // fallback local update
      if (editingId) {
        setRows((rs) => rs.map((r) => (r.id === editingId ? { ...payload } : r)));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, payload]);
      }
      setForm(EMPTY);
    }
  };

  const handleEdit = (row) => {
    setForm(row);
    setEditingId(row.id);
  };

  const handleDelete = async (id) => {
    try {
      const res = await fetch(`${API_BASE}/api/records/${id}`, { method: 'DELETE' });
      if (res.ok || res.status === 204) {
        setRows((rs) => rs.filter((r) => r.id !== id));
      } else {
        throw new Error('Erro ao deletar');
      }
    } catch (err) {
      console.error(err);
      setRows((rs) => rs.filter((r) => r.id !== id));
    }
    if (editingId === id) {
      setEditingId(null);
      setForm(EMPTY);
    }
  };

  const cancelEdit = () => {
    setEditingId(null);
    setForm(EMPTY);
  };

  const handleRuleFileUpload = async (event, ruleKey) => {
    const files = Array.from(event.target.files || []);
    if (files.length === 0) return;
    const file = files[0];
    const fileObj = await new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve({
        name: file.name,
        type: file.type,
        size: file.size,
        content: reader.result,
      });
      reader.readAsDataURL(file);
    });
    setForm((f) => ({
      ...f,
      regrasArquivos: { ...f.regrasArquivos, [ruleKey]: fileObj },
    }));
    event.target.value = "";
  };

  const handleRemoveRuleFile = (ruleKey) => {
    setForm((f) => ({
      ...f,
      regrasArquivos: { ...f.regrasArquivos, [ruleKey]: null },
    }));
  };

  const monthOptions = useMemo(() => {
    const months = new Set();
    rows.forEach((r) => {
      if (r.data && r.data.length >= 7) months.add(r.data.slice(0, 7));
    });
    return Array.from(months).sort();
  }, [rows]);

  const dashboardRows = useMemo(() => {
    if (dashboardMonth === "all") return rows;
    return rows.filter((r) => r.data && r.data.slice(0, 7) === dashboardMonth);
  }, [rows, dashboardMonth]);

  const formatMonthLabel = (value) => {
    if (!value || value === "all") return "Todos os meses";
    const [year, month] = value.split("-");
    return `${month}/${year}`;
  };

  const filtered = useMemo(() => {
    let subset = rows;
    if (monthFilter !== "all") {
      subset = subset.filter((r) => r.data && r.data.slice(0, 7) === monthFilter);
    }
    if (!filter.trim()) return subset;
    const q = filter.toLowerCase();
    return subset.filter((r) =>
      [r.os, r.encarregado, r.regional, r.tipoServico, r.processo]
        .join(" ")
        .toLowerCase()
        .includes(q)
    );
  }, [rows, filter, monthFilter]);

  const exportXlsx = () => {
    const data = filtered.map((r) => {
      const selectedRules = Object.entries(r.regrasOuro || {}).filter(([, value]) => value).map(([key]) => RULES_DE_OURO[key]);
      return {
        "Data": r.data,
        "Nome do Encarregado": r.encarregado,
        "OS": r.os,
        "Tipo de Serviço": r.tipoServico,
        "Regional": r.regional,
        "Observação": r.observacao,
        "Status": r.conformidade,
        "Descrição da Não Conformidade": r.descNaoConformidade,
        "Regras de Ouro": selectedRules.join("; "),
        "Arquivos": r.regrasArquivos ? Object.entries(r.regrasArquivos).filter(([, file]) => file).map(([key, file]) => `${RULES_DE_OURO[key]}: ${file.name}`).join("; ") : "",
        "Processo": r.processo,
        "Matrícula do Eletricista Líder": r.matriculaLider,
        "Matrícula do Eletricista": r.matriculaEletricista,
        "Registro de Foto": r.registroFoto,
      };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 12 }, { wch: 22 }, { wch: 12 }, { wch: 20 }, { wch: 14 },
      { wch: 20 }, { wch: 16 }, { wch: 30 }, { wch: 34 }, { wch: 20 },
      { wch: 30 }, { wch: 24 }, { wch: 20 }, { wch: 26 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Registros");
    const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
    XLSX.writeFile(wb, `registros_os_${suffix}.xlsx`);
  };

  const exportPdf = () => {
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    const margin = 40;
    const lineHeight = 18;
    let y = margin;
    const headers = [
      "Data", "Encarregado", "OS", "Tipo de Serviço", "Regional", "Status", "Não conformidade", "Regras de Ouro"
    ];
    doc.setFontSize(12);
    doc.text("Relatório de Registros", margin, y);
    y += lineHeight * 1.5;
    doc.setFontSize(10);
    doc.text(headers.join("  |  "), margin, y);
    y += lineHeight;
    doc.setLineWidth(0.5);
    doc.line(margin, y - 8, 555, y - 8);

    filtered.forEach((r) => {
      if (y > 740) {
        doc.addPage();
        y = margin;
      }
      const selectedRules = Object.entries(r.regrasOuro || {}).filter(([, value]) => value).map(([key]) => RULES_DE_OURO[key]).join("; ");
      const row = [
        r.data || "-",
        r.encarregado || "-",
        r.os || "-",
        r.tipoServico || "-",
        r.regional || "-",
        r.conformidade || "-",
        r.descNaoConformidade || "-",
        selectedRules || "-",
      ];
      doc.text(row.join("  |  "), margin, y);
      y += lineHeight;
    });
    const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
    doc.save(`registros_os_${suffix}.pdf`);
  };

  const totals = useMemo(() => {
    const conforme = rows.filter((r) => r.conformidade === "Conforme").length;
    const naoConforme = rows.length - conforme;
    return { conforme, naoConforme };
  }, [rows]);

  const dash = useMemo(() => {
    const byRegional = {};
    const byTipo = {};
    const byFoto = { "Enviado": 0, "Enviado com Não conformidade": 0, "Não enviado": 0 };
    const byDate = {};
    const byEncarregado = {};

    dashboardRows.forEach((r) => {
      const reg = r.regional || "Não informado";
      if (!byRegional[reg]) byRegional[reg] = { name: reg, Conforme: 0, "Não conforme": 0 };
      byRegional[reg][r.conformidade] += 1;

      const tipo = r.tipoServico || "Não informado";
      byTipo[tipo] = (byTipo[tipo] || 0) + 1;

      byFoto[r.registroFoto] = (byFoto[r.registroFoto] || 0) + 1;

      if (r.data) {
        if (!byDate[r.data]) byDate[r.data] = { date: r.data, Conforme: 0, "Não conforme": 0 };
        byDate[r.data][r.conformidade] += 1;
      }

      const enc = r.encarregado || "Não informado";
      byEncarregado[enc] = (byEncarregado[enc] || 0) + 1;
    });

    const regionalData = Object.values(byRegional);
    const tipoData = Object.entries(byTipo).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    const fotoData = Object.entries(byFoto).map(([name, value]) => ({ name, value }));
    const dateData = Object.values(byDate).sort((a, b) => a.date.localeCompare(b.date));
    const topEncarregados = Object.entries(byEncarregado).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value).slice(0, 6);

    const conforme = dashboardRows.filter((r) => r.conformidade === "Conforme").length;
    const naoConforme = dashboardRows.length - conforme;
    const taxaConformidade = dashboardRows.length ? Math.round((conforme / dashboardRows.length) * 100) : 0;
    const naoEnviadas = dashboardRows.filter((r) => r.registroFoto === "Não enviado").length;
    const goldRuleCount = dashboardRows.filter((r) => Object.values(r.regrasOuro || {}).some(Boolean)).length;
    const goldRuleRate = dashboardRows.length ? Math.round((goldRuleCount / dashboardRows.length) * 100) : 0;

    return { regionalData, tipoData, fotoData, dateData, topEncarregados, taxaConformidade, naoEnviadas, naoConforme, goldRuleCount, goldRuleRate };
  }, [dashboardRows]);

  const ruleCounts = useMemo(() => {
    const counts = {};
    Object.keys(RULES_DE_OURO).forEach((key) => { counts[key] = 0; });
    rows.forEach((r) => {
      Object.entries(r.regrasOuro || {}).forEach(([key, value]) => {
        if (value && counts[key] !== undefined) counts[key] += 1;
      });
    });
    return counts;
  }, [rows]);

  return (
    <div style={{ minHeight: "100vh", background: "#14181C", fontFamily: "'IBM Plex Mono', ui-monospace, monospace", color: "#E8EBEE", padding: "0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700;800&display=swap');
        .disp { font-family: 'Barlow Condensed', sans-serif; }
        input, select { font-family: 'IBM Plex Mono', monospace; }
        ::placeholder { color: #6B7580; }
        .field-label { font-size: 11px; letter-spacing: 0.08em; text-transform: uppercase; color: #8A93A0; margin-bottom: 6px; display: block; font-weight: 600; }
        .field-input { width: 100%; background: #1C2126; border: 1px solid #2E3540; color: #E8EBEE; padding: 10px 12px; border-radius: 4px; font-size: 14px; outline: none; transition: border-color 0.15s; box-sizing: border-box; }
        .field-input:focus { border-color: #E8930C; }
        .radio-pill { cursor: pointer; padding: 9px 14px; border-radius: 4px; border: 1px solid #2E3540; font-size: 13px; font-weight: 600; text-align: center; transition: all 0.15s; user-select: none; }
        .btn-primary { background: #E8930C; color: #14181C; border: none; padding: 11px 20px; border-radius: 4px; font-weight: 700; font-size: 13px; letter-spacing: 0.03em; text-transform: uppercase; cursor: pointer; display: inline-flex; align-items: center; gap: 8px; transition: filter 0.15s; }
        .btn-primary:hover { filter: brightness(1.1); }
        .btn-primary:disabled { opacity: 0.4; cursor: not-allowed; }
        .btn-ghost { background: transparent; color: #8A93A0; border: 1px solid #2E3540; padding: 11px 16px; border-radius: 4px; font-weight: 600; font-size: 13px; cursor: pointer; }
        .btn-ghost:hover { color: #E8EBEE; border-color: #4A5462; }
        table { border-collapse: collapse; width: 100%; }
        th { text-align: left; font-size: 10px; letter-spacing: 0.07em; text-transform: uppercase; color: #6B7580; padding: 10px 12px; border-bottom: 1px solid #2E3540; font-weight: 600; white-space: nowrap; }
        td { padding: 12px; border-bottom: 1px solid #1E242A; font-size: 13px; vertical-align: middle; }
        tr:hover td { background: #1A1F24; }
      `}</style>

      {/* Header */}
      <div style={{ borderBottom: "1px solid #2E3540", padding: "20px 32px", display: "flex", alignItems: "center", justifyContent: "space-between", background: "#171B20" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 34, height: 34, background: "#E8930C", borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Zap size={19} color="#14181C" strokeWidth={2.5} />
          </div>
          <div>
            <div className="disp" style={{ fontSize: 22, fontWeight: 800, lineHeight: 1 }}>REGISTRO DE OS</div>
            <div style={{ fontSize: 11, color: "#6B7580", letterSpacing: "0.04em" }}>Serviços elétricos · controle de conformidade</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 28 }}>
          <div style={{ display: "flex", gap: 6, background: "#1C2126", padding: 4, borderRadius: 6, border: "1px solid #2E3540" }}>
            <button
              onClick={() => setView("registro")}
              style={{
                display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
                padding: "8px 14px", borderRadius: 4, fontSize: 12, fontWeight: 700, letterSpacing: "0.02em",
                background: view === "registro" ? "#E8930C" : "transparent",
                color: view === "registro" ? "#14181C" : "#8A93A0",
              }}
            >
              <ClipboardList size={14} /> REGISTRO
            </button>
            <button
              onClick={() => setView("dashboard")}
              style={{
                display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
                padding: "8px 14px", borderRadius: 4, fontSize: 12, fontWeight: 700, letterSpacing: "0.02em",
                background: view === "dashboard" ? "#E8930C" : "transparent",
                color: view === "dashboard" ? "#14181C" : "#8A93A0",
              }}
            >
              <LayoutGrid size={14} /> DASHBOARD
            </button>
            <button
              onClick={() => setView("regras")}
              style={{
                display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
                padding: "8px 14px", borderRadius: 4, fontSize: 12, fontWeight: 700, letterSpacing: "0.02em",
                background: view === "regras" ? "#E8930C" : "transparent",
                color: view === "regras" ? "#14181C" : "#8A93A0",
              }}
            >
              <CheckCircle2 size={14} /> REGRAS DE OURO
            </button>
          </div>

          <div style={{ display: "flex", gap: 24 }}>
            <div style={{ textAlign: "right" }}>
              <div className="disp" style={{ fontSize: 26, fontWeight: 800, color: "#E8EBEE", lineHeight: 1 }}>{rows.length}</div>
              <div style={{ fontSize: 10, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.06em" }}>Registros</div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div className="disp" style={{ fontSize: 26, fontWeight: 800, color: "#2F9E52", lineHeight: 1 }}>{totals.conforme}</div>
              <div style={{ fontSize: 10, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.06em" }}>Conformes</div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div className="disp" style={{ fontSize: 26, fontWeight: 800, color: "#D64545", lineHeight: 1 }}>{totals.naoConforme}</div>
              <div style={{ fontSize: 10, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.06em" }}>Não conformes</div>
            </div>
          </div>
        </div>
      </div>

      <div style={{ padding: "28px 32px", maxWidth: 1280, margin: "0 auto" }}>
      {view === "registro" && (
      <>
        {/* Form */}
        <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 24, marginBottom: 28 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
            <div className="disp" style={{ fontSize: 17, fontWeight: 700, letterSpacing: "0.01em" }}>
              {editingId ? "EDITAR REGISTRO" : "NOVO REGISTRO"}
            </div>
            {editingId && (
              <button className="btn-ghost" onClick={cancelEdit}>Cancelar edição</button>
            )}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 16 }}>
            <div>
              <label className="field-label">Data *</label>
              <input type="date" className="field-input" value={form.data} onChange={(e) => update("data", e.target.value)} />
            </div>
            <div>
              <label className="field-label">Nome do Encarregado *</label>
              <input type="text" className="field-input" placeholder="Ex: João Silva" value={form.encarregado} onChange={(e) => update("encarregado", e.target.value)} />
            </div>
            <div>
              <label className="field-label">OS *</label>
              <input type="text" className="field-input" placeholder="Nº da OS" value={form.os} onChange={(e) => update("os", e.target.value)} />
            </div>
            <div>
              <label className="field-label">Tipo de Serviço</label>
              <input type="text" className="field-input" placeholder="Ex: Manutenção preventiva" value={form.tipoServico} onChange={(e) => update("tipoServico", e.target.value)} />
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 16 }}>
            <div>
              <label className="field-label">Regional</label>
              <select className="field-input" value={form.regional} onChange={(e) => update("regional", e.target.value)}>
                <option value="">Selecione</option>
                <option value="Baixada">Baixada</option>
                <option value="Oeste">Oeste</option>
              </select>
            </div>
            <div>
              <label className="field-label">Processo</label>
              <input type="text" className="field-input" placeholder="Ex: Poda / Inspeção" value={form.processo} onChange={(e) => update("processo", e.target.value)} />
            </div>
            <div>
              <label className="field-label">Nome do Eletricista Líder</label>
              <input type="text" className="field-input" placeholder="Ex: João Silva" value={form.nomeEletricistaLider} onChange={(e) => update("nomeEletricistaLider", e.target.value)} />
            </div>
            <div>
              <label className="field-label">Nome do Eletricista</label>
              <input type="text" className="field-input" placeholder="Ex: Pedro Almeida" value={form.nomeEletricista} onChange={(e) => update("nomeEletricista", e.target.value)} />
            </div>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 16 }}>
            <div>
              <label className="field-label">Quem inspecionou</label>
              <select className="field-input" value={form.quemInspecionou} onChange={(e) => update("quemInspecionou", e.target.value)}>
                <option value="">Selecione</option>
                <option value="Cleyton">Cleyton</option>
                <option value="Camilly">Camilly</option>
              </select>
            </div>
            <div>
              <label className="field-label">Observação</label>
              <input type="text" className="field-input" placeholder="Observações adicionais" value={form.observacao} onChange={(e) => update("observacao", e.target.value)} />
            </div>
            <div />
            <div />
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 24, marginBottom: 20 }}>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 24, marginBottom: 20 }}>
            <div>
              <label className="field-label">Conformidade</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                <div
                  className="radio-pill"
                  onClick={() => setForm((f) => ({ ...f, conformidade: "Conforme", descNaoConformidade: "" }))}
                  style={{
                    background: form.conformidade === "Conforme" ? "#1F6B3A" : "#1C2126",
                    borderColor: form.conformidade === "Conforme" ? "#2F9E52" : "#2E3540",
                    color: form.conformidade === "Conforme" ? "#EAF4EE" : "#8A93A0",
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
                  }}
                >
                  <CheckCircle2 size={14} /> Conforme
                </div>
                <div
                  className="radio-pill"
                  onClick={() => update("conformidade", "Não conforme")}
                  style={{
                    background: form.conformidade === "Não conforme" ? "#7A2626" : "#1C2126",
                    borderColor: form.conformidade === "Não conforme" ? "#D64545" : "#2E3540",
                    color: form.conformidade === "Não conforme" ? "#FBEAEA" : "#8A93A0",
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 6,
                  }}
                >
                  <XCircle size={14} /> Não conforme
                </div>
              </div>
              {form.conformidade === "Não conforme" && (
                <div style={{ marginTop: 10 }}>
                  <label className="field-label" style={{ color: "#E8930C" }}>
                    <AlertTriangle size={11} style={{ display: "inline", marginRight: 4, verticalAlign: -1 }} />
                    Descreva a não conformidade *
                  </label>
                  <textarea
                    className="field-input"
                    style={{ resize: "vertical", minHeight: 54, fontFamily: "inherit" }}
                    placeholder="Ex: EPI incompleto, procedimento de bloqueio não seguido..."
                    value={form.descNaoConformidade}
                    onChange={(e) => update("descNaoConformidade", e.target.value)}
                  />
                </div>
              )}
            </div>

            <div>
              <label className="field-label"><Camera size={11} style={{ display: "inline", marginRight: 4, verticalAlign: -1 }} />Registro de Foto</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                {Object.keys(FOTO_STYLES).map((opt) => {
                  const active = form.registroFoto === opt;
                  const s = FOTO_STYLES[opt];
                  return (
                    <div
                      key={opt}
                      className="radio-pill"
                      onClick={() => update("registroFoto", opt)}
                      style={{
                        background: active ? s.bg : "#1C2126",
                        borderColor: active ? s.dot : "#2E3540",
                        color: active ? s.fg : "#8A93A0",
                        fontSize: 12,
                      }}
                    >
                      {opt}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            <button className="btn-primary" disabled={!requiredOk} onClick={handleSubmit}>
              <Plus size={16} /> {editingId ? "Salvar alterações" : "Adicionar registro"}
            </button>
            {!requiredOk && (
              <span style={{ fontSize: 12, color: "#6B7580", alignSelf: "center" }}>
                {form.conformidade === "Não conforme" && !form.descNaoConformidade.trim()
                  ? "Descreva a não conformidade para adicionar."
                  : "Preencha Data, Encarregado e OS para adicionar."}
              </span>
            )}
          </div>
        </div>

        {/* Table */}
        <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, overflow: "hidden" }}>
          <div style={{ padding: "16px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", borderBottom: "1px solid #2E3540" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <input
                  type="text"
                  placeholder="Filtrar por OS, encarregado, regional, processo..."
                  className="field-input"
                  style={{ maxWidth: 340 }}
                  value={filter}
                  onChange={(e) => setFilter(e.target.value)}
                />
                <select
                  value={monthFilter}
                  onChange={(e) => setMonthFilter(e.target.value)}
                  className="field-input"
                  style={{ maxWidth: 180 }}
                >
                  <option value="all">Todos os meses</option>
                  {monthOptions.map((month) => (
                    <option key={month} value={month}>{formatMonthLabel(month)}</option>
                  ))}
                </select>
              </div>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button className="btn-ghost" onClick={exportXlsx} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                  <Download size={14} /> XLS
                </button>
                <button className="btn-ghost" onClick={exportPdf} style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                  <Download size={14} /> PDF
                </button>
              </div>
            </div>
          <div style={{ overflowX: "auto" }}>
            <table>
              <thead>
                <tr>
                  <th>Data</th>
                  <th>Encarregado</th>
                  <th>OS</th>
                  <th>Tipo de Serviço</th>
                  <th>Regional</th>
                  <th>Observação</th>
                  <th>Nome Líder</th>
                  <th>Nome Eletricista</th>
                  <th>Quem inspecionou</th>
                  <th>Processo</th>
                  <th>Arquivos de Regras</th>
                  <th>Regras de Ouro</th>
                  <th>Foto</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr>
                  <td colSpan={14} style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                      {rows.length === 0 ? "Nenhum registro ainda. Preencha o formulário acima para começar." : "Nenhum resultado para esse filtro."}
                    </td>
                  </tr>
                )}
                {filtered.map((r) => {
                  const fs = FOTO_STYLES[r.registroFoto];
                  return (
                    <tr key={r.id}>
                      <td>{r.data}</td>
                      <td>{r.encarregado}</td>
                      <td style={{ color: "#E8930C", fontWeight: 600 }}>{r.os}</td>
                      <td>{r.tipoServico || "—"}</td>
                      <td>{r.regional || "—"}</td>
                      <td>
                        <span
                          title={r.conformidade === "Não conforme" ? r.descNaoConformidade : undefined}
                          style={{
                            fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 20,
                            background: r.conformidade === "Conforme" ? "#1F6B3A" : "#7A2626",
                            color: r.conformidade === "Conforme" ? "#EAF4EE" : "#FBEAEA",
                            cursor: r.conformidade === "Não conforme" ? "help" : "default",
                          }}
                        >
                          {r.conformidade}{r.conformidade === "Não conforme" ? " ⓘ" : ""}
                        </span>
                      </td>
                      <td>{r.processo || "—"}</td>
                      <td>{r.observacao || "—"}</td>
                      <td>{r.nomeEletricistaLider || "—"}</td>
                      <td>{r.nomeEletricista || "—"}</td>
                      <td>{r.quemInspecionou || "—"}</td>
                      <td>
                        <div style={{ display: "grid", gap: 4 }}>
                          {Object.entries(r.regrasArquivos || {}).filter(([, file]) => file).map(([key, file]) => (
                            <span key={`${key}-${file.name}`} style={{ fontSize: 11, color: "#E8EBEE" }}>{`${RULES_DE_OURO[key]}: ${file.name}`}</span>
                          ))}
                        </div>
                      </td>
                      <td>
                        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                          {Object.entries(r.regrasOuro || {})
                            .filter(([, value]) => value)
                            .map(([key]) => (
                              <span key={key} style={{ fontSize: 11, color: "#8A93A0" }}>{RULES_DE_OURO[key]}</span>
                            ))}
                        </div>
                      </td>
                      <td>
                        <span style={{ fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 20, background: fs.bg, color: fs.fg }}>
                          {r.registroFoto}
                        </span>
                      </td>
                      <td>
                        <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                          <button
                            onClick={() => handleEdit(r)}
                            style={{ background: "none", border: "1px solid #2E3540", color: "#8A93A0", borderRadius: 4, padding: "5px 9px", fontSize: 11, cursor: "pointer", fontWeight: 600 }}
                          >
                            Editar
                          </button>
                          <button
                            onClick={() => handleDelete(r.id)}
                            style={{ background: "none", border: "1px solid #2E3540", color: "#D64545", borderRadius: 4, padding: "5px 8px", cursor: "pointer", display: "flex", alignItems: "center" }}
                          >
                            <Trash2 size={13} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </>
      )}

      {view === "dashboard" && (
        <DashboardView
          dash={dash}
          rowsCount={dashboardRows.length}
          monthOptions={monthOptions}
          dashboardMonth={dashboardMonth}
          setDashboardMonth={setDashboardMonth}
        />
      )}
      {view === "regras" && (
        <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 24, marginBottom: 28 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
            <div className="disp" style={{ fontSize: 17, fontWeight: 700, letterSpacing: "0.01em" }}>
              REGRAS DE OURO / NR10
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 12, marginBottom: 20 }}>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 16 }}>
              <label className="field-label" style={{ marginBottom: 10 }}>Encarregado</label>
              <input
                value={form.encarregado}
                onChange={(e) => update("encarregado", e.target.value)}
                className="field-input"
                placeholder="Digite o encarregado"
              />
            </div>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 16 }}>
              <label className="field-label" style={{ marginBottom: 10 }}>OS</label>
              <input
                value={form.os}
                onChange={(e) => update("os", e.target.value)}
                className="field-input"
                placeholder="Digite a OS"
              />
            </div>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 16 }}>
              <label className="field-label" style={{ marginBottom: 10 }}>Tipo de Serviço</label>
              <input
                value={form.tipoServico}
                onChange={(e) => update("tipoServico", e.target.value)}
                className="field-input"
                placeholder="Digite o tipo de serviço"
              />
            </div>
          </div>
          <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 20, marginBottom: 20 }}>
            <div className="disp" style={{ fontSize: 15, fontWeight: 700, marginBottom: 12 }}>Status de conformidade</div>
            <div style={{ display: "grid", gap: 14 }}>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                {['Conforme', 'Não conforme'].map((option) => {
                  const active = form.conformidade === option;
                  return (
                    <div
                      key={option}
                      onClick={() => update('conformidade', option)}
                      style={{
                        cursor: 'pointer', padding: '10px 14px', borderRadius: 6,
                        border: `1px solid ${active ? '#E8930C' : '#2E3540'}`,
                        background: active ? '#E8930C' : '#14181C',
                        color: active ? '#14181C' : '#E8EBEE',
                        fontWeight: 700,
                        minWidth: 140,
                        textAlign: 'center',
                      }}
                    >{option}</div>
                  );
                })}
              </div>
              {form.conformidade === 'Não conforme' && (
                <div style={{ display: 'grid', gap: 8 }}>
                  <div style={{ fontSize: 12, color: '#8A93A0' }}>Descreva a não conformidade</div>
                  <textarea
                    className="field-input"
                    style={{ minHeight: 100, resize: 'vertical', fontFamily: 'inherit' }}
                    value={form.descNaoConformidade}
                    onChange={(e) => update('descNaoConformidade', e.target.value)}
                    placeholder="Descreva a não conformidade para este registro"
                  />
                </div>
              )}
            </div>
          </div>
          <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 20, marginBottom: 20 }}>
            <div className="disp" style={{ fontSize: 15, fontWeight: 700, marginBottom: 12 }}>Marque as etapas aplicáveis ao registro atual</div>
            <div style={{ display: "grid", gap: 12 }}>
              {Object.entries(RULES_DE_OURO).map(([key, label]) => (
                <div key={key} style={{ display: "grid", gridTemplateColumns: "auto 1fr", gap: 14, alignItems: "center", padding: "12px 14px", background: form.regrasOuro[key] ? "#1F2E1F" : "#14181C", borderRadius: 8, border: "1px solid #2E3540" }}>
                  <img src={getRuleImage(key)} alt={label} style={{ width: 56, height: 56, borderRadius: 12, objectFit: "cover", background: "#0F1519" }} />
                  <div style={{ display: "grid", gap: 10 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer", margin: 0 }}>
                      <input
                        type="checkbox"
                        checked={form.regrasOuro[key]}
                        onChange={(e) => update("regrasOuro", { ...form.regrasOuro, [key]: e.target.checked })}
                        style={{ width: 16, height: 16 }}
                      />
                      <span style={{ fontSize: 13, color: "#E8EBEE" }}>{label}</span>
                    </label>
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      <input
                        type="file"
                        className="field-input"
                        onChange={(e) => handleRuleFileUpload(e, key)}
                        style={{ padding: 8 }}
                      />
                      {form.regrasArquivos[key] && (
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 10px", background: "#14181C", border: "1px solid #2E3540", borderRadius: 6 }}>
                          <span style={{ fontSize: 12, color: "#E8EBEE", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{form.regrasArquivos[key].name}</span>
                          <button
                            onClick={() => handleRemoveRuleFile(key)}
                            style={{ background: "none", border: "1px solid #2E3540", color: "#D64545", borderRadius: 4, padding: "4px 8px", cursor: "pointer", fontSize: 12 }}
                          >Remover</button>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 16, marginBottom: 20 }}>
            {Object.entries(RULES_DE_OURO).map(([key, label]) => (
              <div key={key} style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 18 }}>
                <div style={{ fontSize: 13, color: "#E8EBEE", marginBottom: 10, fontWeight: 700 }}>{label}</div>
                <div style={{ fontSize: 12, color: "#8A93A0", marginBottom: 14 }}>Registros com essa regra:</div>
                <div style={{ fontSize: 28, fontWeight: 800, color: "#E8930C" }}>{ruleCounts[key] || 0}</div>
              </div>
            ))}
          </div>
          <div style={{ background: "#14181C", border: "1px solid #2E3540", borderRadius: 8, padding: 18 }}>
            <div style={{ fontSize: 12, color: "#6B7580", marginBottom: 12 }}>A seleção feita aqui será aplicada ao registro atual. Depois volte à aba Registro para salvar ou editar o registro com essas regras.</div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 12 }}>
              {Object.entries(ruleCounts).map(([key, count]) => (
                <div key={key} style={{ padding: 12, background: "#1C2126", borderRadius: 6, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13, color: "#E8EBEE" }}>{RULES_DE_OURO[key]}</span>
                  <span style={{ fontSize: 14, color: "#E8930C", fontWeight: 700 }}>{count}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      </div>
    </div>
  );
}

function DashboardView({ dash, rowsCount, monthOptions, dashboardMonth, setDashboardMonth }) {
  if (rowsCount === 0) {
    return (
      <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: "60px 24px", textAlign: "center", color: "#6B7580" }}>
        Nenhum dado ainda. Adicione registros na aba Registro para ver o dashboard.
      </div>
    );
  }

  const cardStyle = { background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 20 };
  const titleStyle = { fontSize: 12, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", color: "#8A93A0", marginBottom: 16 };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 16, justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ fontSize: 12, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.08em" }}>Filtrar mês</span>
          <select
            value={dashboardMonth}
            onChange={(e) => setDashboardMonth(e.target.value)}
            className="field-input"
            style={{ maxWidth: 180 }}
          >
            <option value="all">Todos os meses</option>
            {monthOptions.map((month) => (
              <option key={month} value={month}>{`${month.slice(5)}/${month.slice(0, 4)}`}</option>
            ))}
          </select>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
          <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 16, minWidth: 170 }}>
            <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 6, textTransform: "uppercase", letterSpacing: "0.08em" }}>Regras de Ouro</div>
            <div className="disp" style={{ fontSize: 26, fontWeight: 800, color: "#E8930C" }}>{dash.goldRuleCount}</div>
            <div style={{ fontSize: 11, color: "#8A93A0" }}>{dash.goldRuleRate}% dos registros</div>
          </div>
        </div>
      </div>
      {/* KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16 }}>
        <div style={cardStyle}>
          <div style={titleStyle}>Taxa de conformidade</div>
          <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: dash.taxaConformidade >= 80 ? "#2F9E52" : dash.taxaConformidade >= 50 ? "#E8930C" : "#D64545" }}>
            {dash.taxaConformidade}%
          </div>
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Total de registros</div>
          <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: "#E8EBEE" }}>{rowsCount}</div>
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Não conformidades</div>
          <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: "#D64545" }}>{dash.naoConforme}</div>
        </div>
        <div style={cardStyle}>
          <div style={titleStyle}>Fotos não enviadas</div>
          <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: dash.naoEnviadas > 0 ? "#E8930C" : "#2F9E52" }}>{dash.naoEnviadas}</div>
        </div>
      </div>

      {/* Regional bar + Foto pie */}
      <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 16 }}>
        <div style={cardStyle}>
          <div style={titleStyle}>Conformidade por regional</div>
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={dash.regionalData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" vertical={false} />
              <XAxis dataKey="name" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} />
              <YAxis stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
              <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} labelStyle={{ color: "#E8EBEE" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="Conforme" stackId="a" fill={CHART_COLORS.conforme} radius={[0, 0, 0, 0]} />
              <Bar dataKey="Não conforme" stackId="a" fill={CHART_COLORS.naoConforme} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={cardStyle}>
          <div style={titleStyle}>Registro de foto</div>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={dash.fotoData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={90} paddingAngle={3}>
                {dash.fotoData.map((entry, i) => {
                  const colors = { "Enviado": CHART_COLORS.conforme, "Enviado com Não conformidade": CHART_COLORS.accent, "Não enviado": CHART_COLORS.naoConforme };
                  return <Cell key={i} fill={colors[entry.name]} />;
                })}
              </Pie>
              <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} />
              <Legend wrapperStyle={{ fontSize: 11 }} layout="vertical" verticalAlign="middle" align="right" />
            </PieChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Timeline + Tipo servico + top encarregados */}
      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 16 }}>
        <div style={cardStyle}>
          <div style={titleStyle}>Evolução por data</div>
          <ResponsiveContainer width="100%" height={240}>
            <LineChart data={dash.dateData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" vertical={false} />
              <XAxis dataKey="date" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} />
              <YAxis stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
              <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} labelStyle={{ color: "#E8EBEE" }} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Line type="monotone" dataKey="Conforme" stroke={CHART_COLORS.conforme} strokeWidth={2} dot={{ r: 3 }} />
              <Line type="monotone" dataKey="Não conforme" stroke={CHART_COLORS.naoConforme} strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div style={cardStyle}>
          <div style={titleStyle}>Top encarregados (nº de OS)</div>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={dash.topEncarregados} layout="vertical" margin={{ left: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" horizontal={false} />
              <XAxis type="number" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
              <YAxis dataKey="name" type="category" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} width={100} />
              <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} />
              <Bar dataKey="value" fill={CHART_COLORS.accent2} radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}
