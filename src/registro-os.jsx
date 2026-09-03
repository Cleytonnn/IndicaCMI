import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import {
  Zap, Plus, Trash2, Download, Camera, CheckCircle2, XCircle,
  LayoutGrid, ClipboardList, AlertTriangle, UserPlus, Film,
  Calendar, Building2, Users, ShieldCheck, FileText, Check,
  Layers, ChevronRight, Edit3, Sparkles, MapPin
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  PieChart, Pie, Cell, LineChart, Line, Legend,
} from "recharts";

const EMPTY = {
  data: "",
  contrato: "LIGHT",
  encarregado: "",
  os: "",
  tipoServico: "",
  regional: "",
  conformidade: "Conforme",
  descNaoConformidade: "",
  processo: "",
  nomeEletricistaLider: "",
  nomeEletricista: "",
  quemInspecionou: "Cleyton",
  tipoInspecao: "",
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
  tipoRegistro: "registro",
};

const EMPTY_NAO_ENVIO = {
  data: "",
  encarregado: "",
  filial: "",
  os: "",
  naoEnvio: "",
  tipoRegistro: "naoEnvio",
};

const EMPTY_MONITORIA = {
  contrato: "LIGHT",
  regional: "",
  dataFilmagem: "",
  placaVeiculo: "",
  fiscal: "",
  setor: "",
  supervisorResponsavel: "",
  eletricistas: [""],
  motoristas: [""],
  houveNaoConformidade: "Não",
  descNaoConformidade: "",
  agenteAgressor: "",
  tratativas: "",
  tipoRegistro: "monitoria",
};

const INSPECTORES_LIGHT = ["Cleyton"];
const INSPECTORES_ENEL = [
  "ALEXANDRO TOLEDO GONÇALVES",
  "ANDERSON AMADURO",
  "GLAUCIA OLIVEIRA",
  "JOSE DOS REIS",
  "DEBORA BRASIL",
];

const FISCAIS_OPTIONS = [
  "ALEXANDRO TOLEDO",
  "GLAUCIA PASSOS",
  "DEBORA BRASIL",
  "JOSÉ JUNIOR",
  "GUILHERME PANTALEÃO",
  "LUCIANA MONTANHA",
  "WESLEY SOUZA",
  "KENNEDY CAMPOS",
  "VINICIUS DUTRA",
  "NIVALDO TEIXEIRA",
  "MARCELLO",
  "BRUNO CESAR",
  "DOMINGOS AGUIAR",
  "MAURICIO",
];

const SETORES_OPTIONS = [
  "NOVAS LIGAÇÃO",
  "COBRANÇA",
  "LINHA VIVA",
  "PODA",
  "EMERGÊNCIA",
  "OBRA",
  "MANUTENÇÃO",
];

const SUPERVISORES_OPTIONS = [
  "WESLEY SOUZA",
  "KENNEDY CAMPOS",
  "VINICIUS DUTRA",
  "NIVALDO TEIXEIRA",
  "MARCELLO",
  "BRUNO CESAR",
  "DOMINGOS AGUIAR",
  "MAURICIO",
];

const AGENTES_AGRESSORES_OPTIONS = [
  "EPI",
  "EPC",
  "FERRAMENTAL",
  "COMPORTAMENTAL",
  "SUPERVISÃO",
  "VEÍCULO",
  "PROCEDIMENTO",
  "5RO",
  "NÃO FILMA ATIVIDADE",
  "DVR-R COM DEFEITO",
  "MOVIMENTAÇÃO DE CARGA",
];

const TRATATIVAS_OPTIONS = [
  "REUNIÃO COM A EQUIPE",
  "DRI",
  "ADVERTÊNCIA",
  "SUSPENSÃO",
  "RECICLAGEM",
  "APLICAR POLITICA DE CONSEQUÊNCIA",
];

const RULES_DE_OURO = {
  desligamentoRede: "Seccionar",
  seccionamento: "Impedir o religamento (bloqueio)",
  bloqueio: "Sinalizar",
  atestar: "Confirmar ausência de tensão",
  protegerEquipamentosEnergizados: "Instalar aterramento temporário (quando aplicável)",
  epi: "Proteger partes energizadas próximas",
};

const RULE_IMAGES = {
  desligamentoRede: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M30 40h60v10H30zM30 60h50v10H30zM30 80h30v10H30z' fill='%23E8930C'/><circle cx='90' cy='70' r='10' fill='%23E8930C'/></svg>`,
  seccionamento: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><path d='M35 30l25 30-25 30' stroke='%23E8930C' stroke-width='10' fill='none'/><path d='M85 30l-25 30 25 30' stroke='%23E8930C' stroke-width='10' fill='none'/></svg>`,
  bloqueio: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23212A30'/><rect x='33' y='55' width='54' height='30' rx='6' fill='%23E8930C'/><path d='M48 55v-15a12 12 0 1 1 24 0v15' stroke='%2314181C' stroke-width='10' fill='none'/><circle cx='89' cy='70' r='5' fill='%2314181C'/></svg>`,
  atestar: `data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 120'><rect width='120' height='120' rx='18' fill='%23E8930C'/><path d='M30 70l20 18 40-52' stroke='%23E8930C' stroke-width='10' fill='none'/></svg>`,
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
  rose: "#F43F5E",
  purple: "#C084FC",
};

const REGIONAIS_POR_CONTRATO = {
  LIGHT: ["Baixada", "Oeste"],
  ENEL: ["Cantagalo", "Macaé", "Pádua", "Angra"],
};

const INSPECAO_POR_CONTRATO = {
  LIGHT: ["Proteção", "RDA"],
  ENEL: ["Obras", "Emergencial", "Comercial"],
};

const PROCESSO_OPTIONS = {
  RDA: [
    "BLINDAGEM (RDA)",
    "EXPANSÃO (RDA)",
    "LINHA VIVA (RDA)",
    "MANOBRA (RDA)",
    "MANUTENÇÃO (RDA)",
    "NORMALIZAÇÃO (RDA)",
    "OPERAÇÃO (RDA)",
  ],
  Proteção: ["CORE", "LIDE", "ANEXO IV", "LNC", "PODA", "REN"],
  Obras: ["OBRAS"],
  Emergencial: ["EMERGENCIAL"],
  Comercial: ["COMERCIAL"],
};

const generateId = () => {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = Math.random() * 16 | 0;
    const v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
};

const getRuleImage = (key) => RULE_IMAGES[key] || "";

const getContrato = (r) => {
  if (r?.contrato) return r.contrato;
  const isEnel = ["Obras", "Emergencial", "Comercial", "OBRAS", "EMERGENCIAL", "COMERCIAL"].includes(r?.tipoInspecao) ||
    ["Cantagalo", "Macaé", "Pádua", "Angra", "CANTAGALO", "MACAÉ", "PADUA", "ANGRA"].includes(r?.regional);
  return isEnel ? "ENEL" : "LIGHT";
};

const compareDateDesc = (a, b) => {
  const da = a?.data || a?.dataFilmagem;
  const db = b?.data || b?.dataFilmagem;
  if (!da && !db) return 0;
  if (!da) return 1;
  if (!db) return -1;
  return new Date(db) - new Date(da);
};

const formatMonthLabel = (value) => {
  if (!value || value === "all") return "Todos os meses";
  const parts = value.split("-");
  if (parts.length === 2) {
    return `${parts[1]}/${parts[0]}`;
  }
  return value;
};

const formatDateForExport = (value) => {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split("-");
    return `${day}/${month}/${year}`;
  }
  return value;
};

const getTodayDate = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
};

const API_BASE = typeof window !== 'undefined' && window.location.hostname === 'localhost' 
  ? 'http://localhost:3000'
  : '';

export default function App() { 
  const [view, setView] = useState("registro"); // "registro" | "monitoria" | "dashboard" | "regras" | "naoEnvio"
  const [dashboardTab, setDashboardTab] = useState("os"); // "os" | "monitoria"
  const [mobileTableView, setMobileTableView] = useState("cards"); // "cards" | "table"
  const [mobileMonitoriaView, setMobileMonitoriaView] = useState("cards"); // "cards" | "table"
  const [rows, setRows] = useState([]);
  const [form, setForm] = useState(EMPTY);
  const [formMonitoria, setFormMonitoria] = useState(EMPTY_MONITORIA);
  const [editingId, setEditingId] = useState(null);
  const [filter, setFilter] = useState("");
  const [monthFilter, setMonthFilter] = useState("all");
  const [dashboardMonth, setDashboardMonth] = useState("all");
  const [dashboardContrato, setDashboardContrato] = useState("all");
  const [dashboardSetor, setDashboardSetor] = useState("all");
  const [currentPage, setCurrentPage] = useState(1);
  const PAGE_SIZE = 20;

  const registroRows = useMemo(() => rows.filter((r) => !r.tipoRegistro || r.tipoRegistro === "registro"), [rows]);
  const naoEnvioRows = useMemo(() => rows.filter((r) => r.tipoRegistro === "naoEnvio"), [rows]);
  const monitoriaRows = useMemo(() => rows.filter((r) => r.tipoRegistro === "monitoria"), [rows]);

  const totals = useMemo(() => {
    let conforme = 0;
    let naoConforme = 0;
    registroRows.forEach((r) => {
      if (r.conformidade === "Conforme") conforme++;
      else if (r.conformidade === "Não conforme") naoConforme++;
    });
    return { conforme, naoConforme };
  }, [registroRows]);

  useEffect(() => {
    const fetchRows = async () => {
      try {
        const res = await fetch(`${API_BASE}/api/records`);
        if (res.ok) {
          const data = await res.json();
          setRows(Array.isArray(data) ? data.slice().sort(compareDateDesc) : []);
          return;
        }
      } catch (e) {
        // fallback to localStorage
      }
      const saved = window.localStorage.getItem("indicaServicosRows");
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed)) setRows(parsed.slice().sort(compareDateDesc));
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
  const updateMonitoria = (field, value) => setFormMonitoria((f) => ({ ...f, [field]: value }));

  const handleContratoChange = (contrato) => {
    setForm((f) => {
      const validInspecoes = INSPECAO_POR_CONTRATO[contrato] || [];
      const isCurrentInspecaoValid = validInspecoes.includes(f.tipoInspecao);
      const validRegionais = REGIONAIS_POR_CONTRATO[contrato] || [];
      const isCurrentRegionalValid = validRegionais.includes(f.regional);
      const defaultQuem = contrato === "LIGHT" ? "Cleyton" : "";

      return {
        ...f,
        contrato,
        regional: isCurrentRegionalValid ? f.regional : "",
        tipoInspecao: isCurrentInspecaoValid ? f.tipoInspecao : "",
        processo: isCurrentInspecaoValid ? f.processo : "",
        quemInspecionou: defaultQuem,
      };
    });
  };

  const handleContratoChangeMonitoria = (contrato) => {
    setFormMonitoria((f) => {
      const validRegionais = REGIONAIS_POR_CONTRATO[contrato] || [];
      const isCurrentRegionalValid = validRegionais.includes(f.regional);
      return {
        ...f,
        contrato,
        regional: isCurrentRegionalValid ? f.regional : "",
      };
    });
  };

  const handleTipoInspecaoChange = (value) => {
    setForm((f) => ({
      ...f,
      tipoInspecao: value,
      processo: f.tipoInspecao === value ? f.processo : "",
    }));
  };

  const setFormToday = () => {
    update("data", getTodayDate());
  };

  const setFormMonitoriaToday = () => {
    updateMonitoria("dataFilmagem", getTodayDate());
  };

  const processoOptions = PROCESSO_OPTIONS[form.tipoInspecao] || [];

  const requiredOk =
    form.data &&
    (form.conformidade !== "Não conforme" || form.descNaoConformidade.trim().length > 0);

  // Monitoria dynamic lists
  const handleAddEletricista = () => {
    setFormMonitoria((f) => ({ ...f, eletricistas: [...f.eletricistas, ""] }));
  };

  const handleRemoveEletricista = (index) => {
    setFormMonitoria((f) => ({
      ...f,
      eletricistas: f.eletricistas.filter((_, i) => i !== index),
    }));
  };

  const handleUpdateEletricista = (index, value) => {
    setFormMonitoria((f) => {
      const next = [...f.eletricistas];
      next[index] = value;
      return { ...f, eletricistas: next };
    });
  };

  const handleAddMotorista = () => {
    setFormMonitoria((f) => ({ ...f, motoristas: [...f.motoristas, ""] }));
  };

  const handleRemoveMotorista = (index) => {
    setFormMonitoria((f) => ({
      ...f,
      motoristas: f.motoristas.filter((_, i) => i !== index),
    }));
  };

  const handleUpdateMotorista = (index, value) => {
    setFormMonitoria((f) => {
      const next = [...f.motoristas];
      next[index] = value;
      return { ...f, motoristas: next };
    });
  };

  const handleSubmit = async () => {
    const isNaoEnvio = form.tipoRegistro === "naoEnvio";
    if (!isNaoEnvio && !requiredOk) return;

    // Se for não envio e a data estiver vazia, preenche automaticamente com a data de hoje
    const finalForm = isNaoEnvio && !form.data ? { ...form, data: getTodayDate() } : form;
    const payload = editingId ? { ...finalForm, id: editingId } : { ...finalForm, id: generateId() };
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
        setRows((rs) => rs.map((r) => (r.id === editingId ? saved : r)).slice().sort(compareDateDesc));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, saved].slice().sort(compareDateDesc));
      }
      setForm(isNaoEnvio ? EMPTY_NAO_ENVIO : EMPTY);
    } catch (err) {
      console.error(err);
      if (editingId) {
        setRows((rs) => rs.map((r) => (r.id === editingId ? { ...payload } : r)).slice().sort(compareDateDesc));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, payload].slice().sort(compareDateDesc));
      }
      setForm(isNaoEnvio ? EMPTY_NAO_ENVIO : EMPTY);
    }
  };

  const handleSubmitMonitoria = async () => {
    if (!formMonitoria.dataFilmagem || !formMonitoria.placaVeiculo) return;
    if (formMonitoria.houveNaoConformidade === "Sim" && !formMonitoria.descNaoConformidade.trim()) return;

    const payload = editingId 
      ? { ...formMonitoria, id: editingId, data: formMonitoria.dataFilmagem, contrato: formMonitoria.contrato || "LIGHT" }
      : { ...formMonitoria, id: generateId(), data: formMonitoria.dataFilmagem, contrato: formMonitoria.contrato || "LIGHT" };

    try {
      const res = await fetch(
        editingId ? `${API_BASE}/api/records/${payload.id}` : `${API_BASE}/api/records`,
        {
          method: editingId ? 'PUT' : 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        }
      );
      if (!res.ok) throw new Error('Falha ao salvar monitoria');
      const saved = await res.json();
      if (editingId) {
        setRows((rs) => rs.map((r) => (r.id === editingId ? saved : r)).slice().sort(compareDateDesc));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, saved].slice().sort(compareDateDesc));
      }
      setFormMonitoria(EMPTY_MONITORIA);
    } catch (err) {
      console.error(err);
      if (editingId) {
        setRows((rs) => rs.map((r) => (r.id === editingId ? { ...payload } : r)).slice().sort(compareDateDesc));
        setEditingId(null);
      } else {
        setRows((rs) => [...rs, payload].slice().sort(compareDateDesc));
      }
      setFormMonitoria(EMPTY_MONITORIA);
    }
  };

  const handleEdit = (row) => {
    if (row.tipoRegistro === "monitoria") {
      const isEnel = row.contrato === "ENEL" ||
        ["Cantagalo", "Macaé", "Pádua", "Angra", "CANTAGALO", "MACAÉ", "PADUA", "ANGRA"].includes(row.regional);
      const inferredContrato = row.contrato || (isEnel ? "ENEL" : "LIGHT");
      setFormMonitoria({
        ...EMPTY_MONITORIA,
        ...row,
        contrato: inferredContrato,
        eletricistas: Array.isArray(row.eletricistas) && row.eletricistas.length > 0 ? row.eletricistas : [""],
        motoristas: Array.isArray(row.motoristas) && row.motoristas.length > 0 ? row.motoristas : [""],
      });
      setEditingId(row.id);
      setView("monitoria");
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }

    if (row.tipoRegistro === "naoEnvio") {
      setForm({
        ...EMPTY_NAO_ENVIO,
        ...row,
      });
      setEditingId(row.id);
      setView("naoEnvio");
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }

    const isEnel = row.contrato === "ENEL" ||
      ["Obras", "Emergencial", "Comercial", "OBRAS", "EMERGENCIAL", "COMERCIAL"].includes(row.tipoInspecao) ||
      ["Cantagalo", "Macaé", "Pádua", "Angra", "CANTAGALO", "MACAÉ", "PADUA", "ANGRA"].includes(row.regional);
    const inferredContrato = row.contrato || (isEnel ? "ENEL" : "LIGHT");
    setForm({
      ...EMPTY,
      ...row,
      contrato: inferredContrato,
    });
    setEditingId(row.id);
    setView("registro");
    window.scrollTo({ top: 0, behavior: 'smooth' });
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
      setForm(view === 'naoEnvio' ? EMPTY_NAO_ENVIO : EMPTY);
      setFormMonitoria(EMPTY_MONITORIA);
    }
  };

  const cancelEdit = () => {
    setEditingId(null);
    setForm(view === 'naoEnvio' ? EMPTY_NAO_ENVIO : EMPTY);
    setFormMonitoria(EMPTY_MONITORIA);
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
    const sourceRows = view === "naoEnvio" 
      ? naoEnvioRows 
      : view === "monitoria" 
      ? monitoriaRows 
      : registroRows;
    sourceRows.forEach((r) => {
      const dt = r.data || r.dataFilmagem;
      if (dt && dt.length >= 7) months.add(dt.slice(0, 7));
    });
    return Array.from(months).sort();
  }, [view, naoEnvioRows, registroRows, monitoriaRows]);

  const dashboardRows = useMemo(() => {
    let list = registroRows;
    if (dashboardMonth !== "all") {
      list = list.filter((r) => r.data && r.data.slice(0, 7) === dashboardMonth);
    }
    if (dashboardContrato !== "all") {
      list = list.filter((r) => getContrato(r) === dashboardContrato);
    }
    return list;
  }, [registroRows, dashboardMonth, dashboardContrato]);

  const monitoriaDashboardRows = useMemo(() => {
    let list = monitoriaRows;
    if (dashboardMonth !== "all") {
      list = list.filter((r) => (r.dataFilmagem || r.data) && (r.dataFilmagem || r.data).slice(0, 7) === dashboardMonth);
    }
    if (dashboardContrato !== "all") {
      list = list.filter((r) => getContrato(r) === dashboardContrato);
    }
    if (dashboardSetor !== "all") {
      list = list.filter((r) => r.setor === dashboardSetor);
    }
    return list;
  }, [monitoriaRows, dashboardMonth, dashboardContrato, dashboardSetor]);

  const contratoStats = useMemo(() => {
    const monthList = dashboardMonth === "all"
      ? registroRows
      : registroRows.filter((r) => r.data && r.data.slice(0, 7) === dashboardMonth);

    let lightCount = 0;
    let enelCount = 0;
    monthList.forEach((r) => {
      if (getContrato(r) === "ENEL") enelCount += 1;
      else lightCount += 1;
    });

    return {
      total: monthList.length,
      light: lightCount,
      enel: enelCount,
    };
  }, [registroRows, dashboardMonth]);

  const contratoStatsMonitoria = useMemo(() => {
    const monthList = dashboardMonth === "all"
      ? monitoriaRows
      : monitoriaRows.filter((r) => (r.dataFilmagem || r.data) && (r.dataFilmagem || r.data).slice(0, 7) === dashboardMonth);

    let lightCount = 0;
    let enelCount = 0;
    monthList.forEach((r) => {
      if (getContrato(r) === "ENEL") enelCount += 1;
      else lightCount += 1;
    });

    return {
      total: monthList.length,
      light: lightCount,
      enel: enelCount,
    };
  }, [monitoriaRows, dashboardMonth]);

  function getWeekKey(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "";
    const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNr = (target.getUTCDay() + 6) % 7;
    target.setUTCDate(target.getUTCDate() - dayNr + 3);
    const firstThursday = new Date(Date.UTC(target.getUTCFullYear(), 0, 4));
    const weekNumber = 1 + Math.round(((target - firstThursday) / 86400000 - 3) / 7);
    return `${target.getUTCFullYear()}-${String(weekNumber).padStart(2, "0")}`;
  }

  const naoEnvioSummary = useMemo(() => {
    const today = new Date().toISOString().slice(0, 10);
    const currentWeek = getWeekKey(today);
    const currentMonth = today.slice(0, 7);
    const uniqueByDay = new Set();
    const uniqueByWeek = new Set();
    const uniqueByMonth = new Set();
    naoEnvioRows.forEach((r) => {
      if (!r.data) return;
      const encarregado = r.encarregado?.trim() || "Não informado";
      if (r.data === today) {
        uniqueByDay.add(encarregado);
      }
      if (getWeekKey(r.data) === currentWeek) {
        uniqueByWeek.add(encarregado);
      }
      if (r.data.slice(0, 7) === currentMonth) {
        uniqueByMonth.add(encarregado);
      }
    });
    return {
      total: naoEnvioRows.length,
      today: uniqueByDay.size,
      week: uniqueByWeek.size,
      month: uniqueByMonth.size,
      currentWeek,
      currentMonth,
      todayLabel: today,
    };
  }, [naoEnvioRows]);

  const filtered = useMemo(() => {
    let subset = view === "naoEnvio" 
      ? naoEnvioRows 
      : view === "monitoria" 
      ? monitoriaRows 
      : registroRows;

    if (monthFilter !== "all") {
      subset = subset.filter((r) => {
        const dt = r.data || r.dataFilmagem;
        return dt && dt.slice(0, 7) === monthFilter;
      });
    }
    if (!filter.trim()) return subset;
    const q = filter.toLowerCase();
    return subset.filter((r) => {
      const fields = [
        r.os, r.encarregado, r.contrato, r.regional, r.tipoServico, r.processo, 
        r.filial, r.naoEnvio, r.tipoRegistro, r.tipoInspecao, r.quemInspecionou,
        r.placaVeiculo, r.fiscal, r.setor, r.supervisorResponsavel, r.agenteAgressor, r.tratativas,
        ...(Array.isArray(r.eletricistas) ? r.eletricistas : []),
        ...(Array.isArray(r.motoristas) ? r.motoristas : [])
      ];
      return fields.filter(Boolean).join(" ").toLowerCase().includes(q);
    });
  }, [view, registroRows, naoEnvioRows, monitoriaRows, filter, monthFilter]);

  useEffect(() => {
    setCurrentPage(1);
  }, [filter, monthFilter, view]);

  const sortedFiltered = useMemo(() => {
    return [...filtered].sort(compareDateDesc);
  }, [filtered]);

  const totalPages = Math.max(1, Math.ceil(sortedFiltered.length / PAGE_SIZE));

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages);
    }
  }, [currentPage, totalPages]);

  const pageRows = useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE;
    return sortedFiltered.slice(start, start + PAGE_SIZE);
  }, [sortedFiltered, currentPage]);

  const ruleCounts = useMemo(() => {
    const counts = {};
    Object.keys(RULES_DE_OURO).forEach((k) => {
      counts[k] = 0;
    });
    registroRows.forEach((r) => {
      if (r.regrasOuro) {
        Object.entries(r.regrasOuro).forEach(([k, v]) => {
          if (v && counts[k] !== undefined) {
            counts[k]++;
          }
        });
      }
    });
    return counts;
  }, [registroRows]);

  const dash = useMemo(() => {
    const total = dashboardRows.length;
    let conformeCount = 0;
    let naoConformeCount = 0;
    let naoEnviadas = 0;
    let goldRuleCount = 0;

    const regionalMap = {};
    const fotoMap = { "Enviado": 0, "Enviado com Não conformidade": 0, "Não enviado": 0 };
    const dateMap = {};
    const encarregadoMap = {};

    let protecaoCount = 0;
    let rdaCount = 0;
    let obrasCount = 0;
    let emergencialCount = 0;
    let comercialCount = 0;

    dashboardRows.forEach((r) => {
      if (r.conformidade === "Conforme") conformeCount++;
      if (r.conformidade === "Não conforme") naoConformeCount++;
      if (r.registroFoto === "Não enviado") naoEnviadas++;

      const hasGold = r.regrasOuro && Object.values(r.regrasOuro).some(Boolean);
      if (hasGold) goldRuleCount++;

      // Inspecao counts
      if (r.tipoInspecao === "Proteção" || r.tipoInspecao === "PROTEÇÃO") protecaoCount++;
      if (r.tipoInspecao === "RDA") rdaCount++;
      if (r.tipoInspecao === "Obras" || r.tipoInspecao === "OBRAS") obrasCount++;
      if (r.tipoInspecao === "Emergencial" || r.tipoInspecao === "EMERGENCIAL") emergencialCount++;
      if (r.tipoInspecao === "Comercial" || r.tipoInspecao === "COMERCIAL") comercialCount++;

      // Regional
      const reg = r.regional || "Não informado";
      if (!regionalMap[reg]) regionalMap[reg] = { name: reg, Conforme: 0, "Não conforme": 0 };
      if (r.conformidade === "Conforme") regionalMap[reg].Conforme++;
      else if (r.conformidade === "Não conforme") regionalMap[reg]["Não conforme"]++;

      // Foto
      if (r.registroFoto && fotoMap[r.registroFoto] !== undefined) {
        fotoMap[r.registroFoto]++;
      }

      // Date
      if (r.data) {
        if (!dateMap[r.data]) dateMap[r.data] = { date: r.data.slice(5), Conforme: 0, "Não conforme": 0 };
        if (r.conformidade === "Conforme") dateMap[r.data].Conforme++;
        else if (r.conformidade === "Não conforme") dateMap[r.data]["Não conforme"]++;
      }

      // Encarregado
      const enc = r.encarregado?.trim();
      if (enc) {
        encarregadoMap[enc] = (encarregadoMap[enc] || 0) + 1;
      }
    });

    const taxaConformidade = total > 0 ? Math.round((conformeCount / total) * 100) : 100;
    const goldRuleRate = total > 0 ? Math.round((goldRuleCount / total) * 100) : 0;

    const regionalData = Object.values(regionalMap);
    const fotoData = Object.entries(fotoMap).map(([name, value]) => ({ name, value }));
    const dateData = Object.entries(dateMap)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([, v]) => v);
    const topEncarregados = Object.entries(encarregadoMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 6);

    return {
      taxaConformidade,
      naoConforme: naoConformeCount,
      naoEnviadas,
      goldRuleCount,
      goldRuleRate,
      protecaoCount,
      rdaCount,
      obrasCount,
      emergencialCount,
      comercialCount,
      regionalData,
      fotoData,
      dateData,
      topEncarregados,
    };
  }, [dashboardRows]);

  const dashMonitoria = useMemo(() => {
    const total = monitoriaDashboardRows.length;
    let conformeCount = 0;
    let naoConformeCount = 0;

    const setorMap = {};
    const regionalMap = {};
    const agenteMap = {};
    const tratativaMap = {};
    const fiscalMap = {};

    monitoriaDashboardRows.forEach((r) => {
      const isNaoConforme = r.houveNaoConformidade === "Sim";
      if (isNaoConforme) naoConformeCount++;
      else conformeCount++;

      // Setor
      const st = r.setor || "Não informado";
      if (!setorMap[st]) setorMap[st] = { name: st, Conforme: 0, "Não conforme": 0 };
      if (isNaoConforme) setorMap[st]["Não conforme"]++;
      else setorMap[st].Conforme++;

      // Regional
      const reg = r.regional || "Não informado";
      if (!regionalMap[reg]) regionalMap[reg] = { name: reg, Conforme: 0, "Não conforme": 0 };
      if (isNaoConforme) regionalMap[reg]["Não conforme"]++;
      else regionalMap[reg].Conforme++;

      // Agente Agressor
      if (r.agenteAgressor) {
        agenteMap[r.agenteAgressor] = (agenteMap[r.agenteAgressor] || 0) + 1;
      }

      // Tratativas
      if (r.tratativas) {
        tratativaMap[r.tratativas] = (tratativaMap[r.tratativas] || 0) + 1;
      }

      // Fiscal
      const fisc = r.fiscal?.trim();
      if (fisc) {
        fiscalMap[fisc] = (fiscalMap[fisc] || 0) + 1;
      }
    });

    const taxaConformidade = total > 0 ? Math.round((conformeCount / total) * 100) : 100;
    const setorData = Object.values(setorMap);
    const regionalData = Object.values(regionalMap);
    const agenteData = Object.entries(agenteMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
    const tratativaData = Object.entries(tratativaMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);
    const fiscalData = Object.entries(fiscalMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8);

    return {
      total,
      conformeCount,
      naoConformeCount,
      taxaConformidade,
      setorData,
      regionalData,
      agenteData,
      tratativaData,
      fiscalData,
    };
  }, [monitoriaDashboardRows]);

  const exportXlsx = () => {
    if (view === 'naoEnvio') {
      const data = filtered.map((r) => ({
        "Data": formatDateForExport(r.data),
        "Nome do Encarregado": r.encarregado || "",
        "Filial": r.filial || "",
        "OS": r.os || "",
        "Motivo do Não Envio": r.naoEnvio || "",
      }));
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [
        { wch: 14 }, { wch: 24 }, { wch: 18 }, { wch: 16 }, { wch: 40 }
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "NaoEnvio");
      const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
      XLSX.writeFile(wb, `nao_envio_${suffix}.xlsx`);
      return;
    }

    if (view === 'monitoria') {
      const data = filtered.map((r) => ({
        "Data da Filmagem": formatDateForExport(r.dataFilmagem || r.data),
        "Contrato": getContrato(r),
        "Regional": r.regional || "",
        "Placa do Veículo": r.placaVeiculo || "",
        "Fiscal": r.fiscal || "",
        "Setor": r.setor || "",
        "Supervisor Responsável": r.supervisorResponsavel || "",
        "Eletricistas": Array.isArray(r.eletricistas) ? r.eletricistas.filter(Boolean).join(", ") : "",
        "Motoristas": Array.isArray(r.motoristas) ? r.motoristas.filter(Boolean).join(", ") : "",
        "Houve Não Conformidade": r.houveNaoConformidade || "Não",
        "Descrição da Não Conformidade": r.descNaoConformidade || "",
        "Agente Agressor": r.agenteAgressor || "",
        "Tratativas": r.tratativas || "",
      }));
      const ws = XLSX.utils.json_to_sheet(data);
      ws["!cols"] = [
        { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 14 }, { wch: 22 }, { wch: 18 }, { wch: 22 },
        { wch: 30 }, { wch: 24 }, { wch: 20 }, { wch: 32 }, { wch: 24 }, { wch: 28 }
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Monitoria");
      const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
      XLSX.writeFile(wb, `monitoria_${suffix}.xlsx`);
      return;
    }

    const data = filtered.map((r) => {
      const rowData = {
        "Data": formatDateForExport(r.data),
        "Contrato": getContrato(r),
        "Encarregado": r.encarregado || "",
        "OS": r.os || "",
        "Tipo de Serviço": r.tipoServico || "",
        "Regional": r.regional || "",
        "Status de Conformidade": r.conformidade || "",
        "Descrição da Não Conformidade": r.descNaoConformidade || "",
        "Observação": r.observacao || "",
        "Nome do Eletricista Líder": r.nomeEletricistaLider || "",
        "Nome do Eletricista": r.nomeEletricista || "",
        "Quem Inspecionou": r.quemInspecionou || "",
        "Tipo de Inspeção": r.tipoInspecao || "",
        "Processo": r.processo || "",
        "Registro de Foto": r.registroFoto || "",
      };
      Object.entries(RULES_DE_OURO).forEach(([key, label]) => {
        rowData[label] = r.regrasOuro?.[key] ? "Sim" : "Não";
      });
      return rowData;
    });

    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [
      { wch: 12 }, { wch: 10 }, { wch: 22 }, { wch: 12 }, { wch: 26 }, { wch: 14 },
      { wch: 22 }, { wch: 32 }, { wch: 22 }, { wch: 24 }, { wch: 24 }, { wch: 20 },
      { wch: 18 }, { wch: 16 }, { wch: 16 }, ...Object.keys(RULES_DE_OURO).map(() => ({ wch: 20 }))
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Registros");
    const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
    XLSX.writeFile(wb, `inspecoes_${suffix}.xlsx`);
  };

  const exportPdf = () => {
    const doc = new jsPDF({ orientation: "landscape" });
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    
    if (view === "naoEnvio") {
      doc.text("Relatório de Não Envio de RO", 14, 15);
      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.text(`Total: ${filtered.length} registros | Mês: ${formatMonthLabel(monthFilter)}`, 14, 21);

      const tableColumn = ["Data", "Encarregado", "Filial", "OS", "Motivo do Não Envio"];
      const tableRows = filtered.map((r) => [
        formatDateForExport(r.data),
        r.encarregado || "-",
        r.filial || "-",
        r.os || "-",
        r.naoEnvio || "-",
      ]);

      autoTable(doc, {
        head: [tableColumn],
        body: tableRows,
        startY: 25,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [232, 147, 12], textColor: 20 },
      });
      const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
      doc.save(`nao_envio_${suffix}.pdf`);
      return;
    }

    if (view === "monitoria") {
      doc.text("Relatório de Monitoria de Filmagem", 14, 15);
      doc.setFontSize(9);
      doc.setFont("helvetica", "normal");
      doc.text(`Total: ${filtered.length} registros | Mês: ${formatMonthLabel(monthFilter)}`, 14, 21);

      const tableColumn = [
        "Data", "Contrato", "Regional", "Placa", "Fiscal", "Setor", "Supervisor",
        "Eletricistas", "Motoristas", "Conformidade", "Agente Agressor", "Tratativas"
      ];
      const tableRows = filtered.map((r) => [
        formatDateForExport(r.dataFilmagem || r.data),
        getContrato(r),
        r.regional || "-",
        r.placaVeiculo || "-",
        r.fiscal || "-",
        r.setor || "-",
        r.supervisorResponsavel || "-",
        Array.isArray(r.eletricistas) ? r.eletricistas.filter(Boolean).join(", ") : "-",
        Array.isArray(r.motoristas) ? r.motoristas.filter(Boolean).join(", ") : "-",
        r.houveNaoConformidade === "Sim" ? "Não Conforme" : "Conforme",
        r.agenteAgressor || "-",
        r.tratativas || "-",
      ]);

      autoTable(doc, {
        head: [tableColumn],
        body: tableRows,
        startY: 25,
        styles: { fontSize: 8 },
        headStyles: { fillColor: [232, 147, 12], textColor: 20 },
      });
      const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
      doc.save(`monitoria_${suffix}.pdf`);
      return;
    }

    doc.text("Relatório de Inspeções de OS", 14, 15);
    doc.setFontSize(9);
    doc.setFont("helvetica", "normal");
    doc.text(`Total: ${filtered.length} registros | Mês: ${formatMonthLabel(monthFilter)}`, 14, 21);

    const tableColumn = [
      "Data", "Contrato", "Encarregado", "OS", "Regional", "Status",
      "Processo", "Quem Insp.", "Tipo Insp.", "Foto"
    ];
    const tableRows = filtered.map((r) => [
      formatDateForExport(r.data),
      getContrato(r),
      r.encarregado || "-",
      r.os || "-",
      r.regional || "-",
      r.conformidade || "-",
      r.processo || "-",
      r.quemInspecionou || "-",
      r.tipoInspecao || "-",
      r.registroFoto || "-",
    ]);

    autoTable(doc, {
      head: [tableColumn],
      body: tableRows,
      startY: 25,
      styles: { fontSize: 8 },
      headStyles: { fillColor: [232, 147, 12], textColor: 20 },
    });

    const suffix = monthFilter === "all" ? new Date().toISOString().slice(0, 10) : `${monthFilter}-historic`;
    doc.save(`inspecoes_${suffix}.pdf`);
  };

  return (
    <div style={{ minHeight: "100vh", background: "#14181C", fontFamily: "'IBM Plex Mono', ui-monospace, monospace", color: "#E8EBEE", padding: "0" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=Barlow+Condensed:wght@600;700;800&display=swap');
        * { box-sizing: border-box; }
        .disp { font-family: 'Barlow Condensed', sans-serif; }
        input, select, textarea { font-family: 'IBM Plex Mono', monospace; }
        ::placeholder { color: #6B7580; }

        .field-label {
          font-size: 11px;
          letter-spacing: 0.08em;
          text-transform: uppercase;
          color: #9AA4B2;
          margin-bottom: 6px;
          display: flex;
          align-items: center;
          gap: 6px;
          font-weight: 600;
        }

        .field-input {
          width: 100%;
          background: #1C2126;
          border: 1px solid #2E3540;
          color: #E8EBEE;
          padding: 11px 12px;
          border-radius: 8px;
          font-size: 13px;
          outline: none;
          transition: all 0.15s ease;
          box-sizing: border-box;
          min-height: 44px;
        }

        .field-input:focus {
          border-color: #E8930C;
          background: #20262D;
          box-shadow: 0 0 0 2px rgba(232, 147, 12, 0.2);
        }

        .radio-pill {
          cursor: pointer;
          padding: 10px 12px;
          border-radius: 8px;
          border: 1px solid #2E3540;
          font-size: 12px;
          font-weight: 600;
          text-align: center;
          transition: all 0.15s ease;
          user-select: none;
          display: flex;
          align-items: center;
          justify-content: center;
          min-height: 44px;
          box-sizing: border-box;
          touch-action: manipulation;
        }

        .radio-pill:hover {
          border-color: #4A5462;
        }

        .radio-pill:active {
          transform: scale(0.98);
        }

        .btn-primary {
          background: #E8930C;
          color: #14181C;
          border: none;
          padding: 12px 22px;
          border-radius: 8px;
          font-weight: 800;
          font-size: 14px;
          letter-spacing: 0.04em;
          text-transform: uppercase;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          transition: all 0.15s ease;
          min-height: 46px;
          touch-action: manipulation;
        }

        .btn-primary:hover {
          filter: brightness(1.1);
          transform: translateY(-1px);
        }

        .btn-primary:active {
          transform: translateY(0) scale(0.99);
        }

        .btn-primary:disabled {
          opacity: 0.4;
          cursor: not-allowed;
          transform: none;
        }

        .btn-ghost {
          background: transparent;
          color: #8A93A0;
          border: 1px solid #2E3540;
          padding: 10px 16px;
          border-radius: 8px;
          font-weight: 600;
          font-size: 13px;
          cursor: pointer;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          gap: 6px;
          min-height: 42px;
          touch-action: manipulation;
        }

        .btn-ghost:hover {
          color: #E8EBEE;
          border-color: #4A5462;
          background: #1C2126;
        }

        .btn-ghost:disabled {
          opacity: 0.4;
          cursor: not-allowed;
        }

        .btn-quick-today {
          background: #252D37;
          border: 1px solid #3A4452;
          color: #38BDF8;
          font-size: 11px;
          font-weight: 700;
          padding: 0 10px;
          border-radius: 6px;
          cursor: pointer;
          height: 28px;
          display: inline-flex;
          align-items: center;
          gap: 4px;
          transition: all 0.15s ease;
        }
        .btn-quick-today:hover {
          background: #38BDF8;
          color: #14181C;
        }

        /* Form Sub-Sections */
        .form-section-card {
          background: #191E24;
          border: 1px solid #28303B;
          border-radius: 10px;
          padding: 18px;
          margin-bottom: 18px;
          transition: border-color 0.15s ease;
        }

        .form-section-card:hover {
          border-color: #384252;
        }

        .form-section-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 8px;
          margin-bottom: 14px;
          padding-bottom: 10px;
          border-bottom: 1px solid #28303B;
        }

        .form-section-title {
          font-family: 'Barlow Condensed', sans-serif;
          font-size: 15px;
          font-weight: 700;
          letter-spacing: 0.05em;
          text-transform: uppercase;
          color: #E8930C;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        /* Layout containers */
        .main-wrapper {
          padding: 24px 32px;
          max-width: 1280px;
          margin: 0 auto;
        }

        .card-box {
          background: #171B20;
          border: 1px solid #2E3540;
          border-radius: 12px;
          padding: 24px;
          margin-bottom: 24px;
          box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
        }

        .app-header {
          border-bottom: 1px solid #2E3540;
          padding: 16px 32px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          background: #171B20;
          gap: 16px;
          flex-wrap: wrap;
        }

        .header-nav {
          display: flex;
          gap: 6px;
          background: #1C2126;
          padding: 5px;
          border-radius: 8px;
          border: 1px solid #2E3540;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
          scrollbar-width: none;
          max-width: 100%;
        }
        .header-nav::-webkit-scrollbar { display: none; }

        .nav-tab-btn {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 6px;
          border: none;
          cursor: pointer;
          padding: 9px 15px;
          border-radius: 6px;
          font-size: 12px;
          font-weight: 700;
          letter-spacing: 0.02em;
          white-space: nowrap;
          transition: all 0.15s ease;
          min-height: 38px;
          touch-action: manipulation;
        }

        .header-stats {
          display: flex;
          gap: 20px;
          flex-wrap: wrap;
        }

        /* Grids */
        .grid-4 {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 16px;
          margin-bottom: 16px;
        }

        .grid-3 {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 16px;
          margin-bottom: 16px;
        }

        .grid-2 {
          display: grid;
          grid-template-columns: 1fr 1.4fr;
          gap: 20px;
          margin-bottom: 20px;
        }

        .grid-kpi {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 16px;
        }

        .grid-charts-main {
          display: grid;
          grid-template-columns: 1.6fr 1fr;
          gap: 16px;
        }

        .grid-charts-sub {
          display: grid;
          grid-template-columns: 1.4fr 1fr;
          gap: 16px;
        }

        .table-container {
          background: #171B20;
          border: 1px solid #2E3540;
          border-radius: 12px;
          overflow: hidden;
          box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
        }

        .table-header-bar {
          padding: 16px 20px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          border-bottom: 1px solid #2E3540;
          gap: 12px;
          flex-wrap: wrap;
        }

        table {
          border-collapse: collapse;
          width: 100%;
        }

        th {
          text-align: left;
          font-size: 11px;
          letter-spacing: 0.07em;
          text-transform: uppercase;
          color: #8A93A0;
          padding: 12px 14px;
          border-bottom: 1px solid #2E3540;
          font-weight: 600;
          white-space: nowrap;
          background: #191E23;
        }

        td {
          padding: 12px 14px;
          border-bottom: 1px solid #1E242A;
          font-size: 13px;
          vertical-align: middle;
        }

        tr:hover td {
          background: #1A1F24;
        }

        /* Mobile Card View for Records */
        .record-card-mobile {
          background: #191E24;
          border: 1px solid #28303B;
          border-radius: 10px;
          padding: 14px 16px;
          margin-bottom: 12px;
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        /* Responsive Breakpoints */
        @media (max-width: 1024px) {
          .grid-4 { grid-template-columns: repeat(2, 1fr); }
          .grid-3 { grid-template-columns: repeat(2, 1fr); }
          .grid-kpi { grid-template-columns: repeat(2, 1fr); }
          .grid-charts-main,
          .grid-charts-sub { grid-template-columns: 1fr; }
        }

        @media (max-width: 768px) {
          /* Prevent auto-zoom in iOS/Android inputs */
          input, select, textarea {
            font-size: 16px !important;
          }

          .app-header {
            padding: 12px 16px;
            flex-direction: column;
            align-items: stretch;
            gap: 12px;
          }
          .header-nav {
            width: 100%;
            overflow-x: auto;
            padding: 4px;
          }
          .header-stats {
            justify-content: space-around;
            width: 100%;
            background: #191E23;
            padding: 8px 12px;
            border-radius: 8px;
            border: 1px solid #28303B;
          }
          .header-stats > div {
            text-align: center !important;
          }
          .main-wrapper {
            padding: 12px 10px;
          }
          .card-box {
            padding: 14px 12px;
            margin-bottom: 16px;
            border-radius: 10px;
          }
          .form-section-card {
            padding: 14px 12px;
            margin-bottom: 14px;
          }
          .grid-4,
          .grid-3,
          .grid-2 {
            grid-template-columns: 1fr;
            gap: 14px;
          }
          .grid-kpi {
            grid-template-columns: repeat(2, 1fr);
            gap: 10px;
          }
          .table-header-bar {
            flex-direction: column;
            align-items: stretch;
            padding: 12px;
          }
          .table-header-bar input,
          .table-header-bar select {
            max-width: 100% !important;
            width: 100%;
          }
          .btn-primary {
            width: 100%;
            padding: 14px;
            font-size: 15px;
          }
        }

        @media (max-width: 480px) {
          .grid-kpi { grid-template-columns: 1fr; }
          .header-stats { gap: 10px; }
          .header-stats .disp { font-size: 18px !important; }
        }
      `}</style>

      {/* Header */}
      <div className="app-header">
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 38, height: 38, background: "#E8930C", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 2px 10px rgba(232, 147, 12, 0.35)" }}>
            <Zap size={22} color="#14181C" strokeWidth={2.5} />
          </div>
          <div>
            <div className="disp" style={{ fontSize: 22, fontWeight: 800, lineHeight: 1 }}>REGISTRO DE OS & MONITORIA</div>
            <div style={{ fontSize: 11, color: "#8A93A0", letterSpacing: "0.04em", marginTop: 2 }}>Serviços elétricos · controle de conformidade</div>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap" }}>
          <div className="header-nav">
            <button
              onClick={() => { setView("registro"); setEditingId(null); }}
              className="nav-tab-btn"
              style={{
                background: view === "registro" ? "#E8930C" : "transparent",
                color: view === "registro" ? "#14181C" : "#8A93A0",
              }}
            >
              <ClipboardList size={14} /> REGISTRO
            </button>
            <button
              onClick={() => { setView("monitoria"); setEditingId(null); }}
              className="nav-tab-btn"
              style={{
                background: view === "monitoria" ? "#E8930C" : "transparent",
                color: view === "monitoria" ? "#14181C" : "#8A93A0",
              }}
            >
              <Film size={14} /> MONITORIA
            </button>
            <button
              onClick={() => setView("dashboard")}
              className="nav-tab-btn"
              style={{
                background: view === "dashboard" ? "#E8930C" : "transparent",
                color: view === "dashboard" ? "#14181C" : "#8A93A0",
              }}
            >
              <LayoutGrid size={14} /> DASHBOARD
            </button>
            <button
              onClick={() => setView("regras")}
              className="nav-tab-btn"
              style={{
                background: view === "regras" ? "#E8930C" : "transparent",
                color: view === "regras" ? "#14181C" : "#8A93A0",
              }}
            >
              <CheckCircle2 size={14} /> REGRAS DE OURO
            </button>
            <button
              onClick={() => {
                setView("naoEnvio");
                setForm(EMPTY_NAO_ENVIO);
                setEditingId(null);
              }}
              className="nav-tab-btn"
              style={{
                background: view === "naoEnvio" ? "#E8930C" : "transparent",
                color: view === "naoEnvio" ? "#14181C" : "#8A93A0",
              }}
            >
              <AlertTriangle size={14} /> NÃO ENVIO
            </button>
          </div>

          <div className="header-stats">
            <div>
              <div className="disp" style={{ fontSize: 22, fontWeight: 800, color: "#E8EBEE", lineHeight: 1 }}>{rows.length}</div>
              <div style={{ fontSize: 10, color: "#8A93A0", textTransform: "uppercase", letterSpacing: "0.06em", marginTop: 2 }}>Registros</div>
            </div>
            <div>
              <div className="disp" style={{ fontSize: 22, fontWeight: 800, color: "#2F9E52", lineHeight: 1 }}>{totals.conforme}</div>
              <div style={{ fontSize: 10, color: "#8A93A0", textTransform: "uppercase", letterSpacing: "0.06em", marginTop: 2 }}>Conformes</div>
            </div>
            <div>
              <div className="disp" style={{ fontSize: 22, fontWeight: 800, color: "#D64545", lineHeight: 1 }}>{totals.naoConforme}</div>
              <div style={{ fontSize: 10, color: "#8A93A0", textTransform: "uppercase", letterSpacing: "0.06em", marginTop: 2 }}>Não conformes</div>
            </div>
          </div>
        </div>
      </div>

      <div className="main-wrapper">
      {view === "registro" && (
      <>
        {/* Form */}
        <div className="card-box">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 8, height: 24, background: "#E8930C", borderRadius: 4 }}></div>
              <div className="disp" style={{ fontSize: 20, fontWeight: 800, letterSpacing: "0.02em" }}>
                {editingId ? "EDITAR REGISTRO DE OS" : "NOVO REGISTRO DE OS"}
              </div>
            </div>
            {editingId && (
              <button className="btn-ghost" onClick={cancelEdit} style={{ height: 36, padding: "0 14px", fontSize: 12 }}>
                Cancelar edição
              </button>
            )}
          </div>

          {/* SEÇÃO 1: CONTRATO & DADOS DA OS */}
          <div className="form-section-card">
            <div className="form-section-header">
              <div className="form-section-title">
                <Building2 size={16} /> 1. Contrato & Dados Principais
              </div>
            </div>

            {/* Contrato Toggle */}
            <div style={{ marginBottom: 16 }}>
              <label className="field-label">Selecione o Contrato *</label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                {['LIGHT', 'ENEL'].map((opt) => {
                  const active = form.contrato === opt;
                  const isLight = opt === 'LIGHT';
                  return (
                    <div
                      key={opt}
                      className="radio-pill"
                      onClick={() => handleContratoChange(opt)}
                      style={{
                        background: active 
                          ? (isLight ? "rgba(56, 189, 248, 0.18)" : "rgba(192, 132, 252, 0.18)") 
                          : "#1C2126",
                        borderColor: active 
                          ? (isLight ? "#38BDF8" : "#C084FC") 
                          : "#2E3540",
                        color: active 
                          ? (isLight ? "#38BDF8" : "#C084FC") 
                          : "#8A93A0",
                        fontSize: 14,
                        fontWeight: 800,
                        gap: 8,
                        boxShadow: active ? `0 0 12px ${isLight ? 'rgba(56, 189, 248, 0.25)' : 'rgba(192, 132, 252, 0.25)'}` : "none",
                      }}
                    >
                      {active && <Check size={16} />}
                      {opt}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Tipo de Inspeção Chips */}
            <div style={{ marginBottom: 16 }}>
              <label className="field-label">Tipo de Inspeção *</label>
              <div style={{ display: "grid", gridTemplateColumns: `repeat(${((INSPECAO_POR_CONTRATO[form.contrato] || []).length) || 1}, 1fr)`, gap: 8 }}>
                {(INSPECAO_POR_CONTRATO[form.contrato] || []).map((opt) => {
                  const active = form.tipoInspecao === opt;
                  return (
                    <div
                      key={opt}
                      className="radio-pill"
                      onClick={() => handleTipoInspecaoChange(opt)}
                      style={{
                        background: active ? "#1F6B3A" : "#1C2126",
                        borderColor: active ? "#2F9E52" : "#2E3540",
                        color: active ? "#EAF4EE" : "#8A93A0",
                        fontSize: 13,
                        fontWeight: 700,
                        padding: "8px 6px",
                      }}
                    >
                      {opt}
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="grid-3">
              {/* Data com botão "Hoje" */}
              <div>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                  <label className="field-label" style={{ margin: 0 }}>Data da Inspeção *</label>
                  <button type="button" onClick={setFormToday} className="btn-quick-today">
                    <Calendar size={12} /> Hoje
                  </button>
                </div>
                <input
                  type="date"
                  className="field-input"
                  value={form.data}
                  onChange={(e) => update("data", e.target.value)}
                />
              </div>

              {/* Regional Chips */}
              <div>
                <label className="field-label">Regional</label>
                <div style={{ display: "grid", gridTemplateColumns: `repeat(${(REGIONAIS_POR_CONTRATO[form.contrato] || []).length || 2}, 1fr)`, gap: 6 }}>
                  {(REGIONAIS_POR_CONTRATO[form.contrato] || []).map((opt) => {
                    const active = form.regional === opt;
                    return (
                      <div
                        key={opt}
                        className="radio-pill"
                        onClick={() => update("regional", opt)}
                        style={{
                          background: active ? "#1F6B3A" : "#1C2126",
                          borderColor: active ? "#2F9E52" : "#2E3540",
                          color: active ? "#EAF4EE" : "#8A93A0",
                          fontSize: 12,
                          padding: "6px 4px",
                          minHeight: 40,
                        }}
                      >
                        {opt}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Número da OS */}
              <div>
                <label className="field-label">Número da OS</label>
                <input
                  type="text"
                  className="field-input"
                  placeholder="Ex: 1234567"
                  value={form.os}
                  onChange={(e) => update("os", e.target.value)}
                  style={{ fontWeight: 600, color: "#E8930C" }}
                />
              </div>
            </div>

            <div className="grid-2" style={{ gridTemplateColumns: "1fr 1fr", margin: 0 }}>
              {/* Processo */}
              <div>
                <label className="field-label">Processo</label>
                <select
                  className="field-input"
                  value={form.processo}
                  onChange={(e) => update("processo", e.target.value)}
                  style={{ appearance: "auto" }}
                  disabled={!form.tipoInspecao}
                >
                  <option value="">{form.tipoInspecao ? "Selecione o processo..." : "Selecione o tipo de inspeção antes"}</option>
                  {processoOptions.map((opt) => (
                    <option key={opt} value={opt}>{opt}</option>
                  ))}
                </select>
              </div>

              {/* Tipo de Serviço */}
              <div>
                <label className="field-label">Tipo de Serviço</label>
                <input
                  type="text"
                  className="field-input"
                  placeholder="Ex: Manutenção preventiva, troca de poste..."
                  value={form.tipoServico}
                  onChange={(e) => update("tipoServico", e.target.value)}
                />
              </div>
            </div>
          </div>

          {/* SEÇÃO 2: EQUIPE & RESPONSÁVEIS */}
          <div className="form-section-card">
            <div className="form-section-header">
              <div className="form-section-title">
                <Users size={16} /> 2. Equipe & Responsáveis
              </div>
            </div>

            <div className="grid-4">
              {/* Quem inspecionou */}
              <div>
                <label className="field-label" style={{ color: form.contrato === "ENEL" ? "#C084FC" : "#38BDF8" }}>
                  {form.contrato === "ENEL" ? "Téc. de Segurança do Trabalho *" : "Quem inspecionou"}
                </label>
                {form.contrato === "ENEL" ? (
                  <select
                    className="field-input"
                    value={form.quemInspecionou}
                    onChange={(e) => update("quemInspecionou", e.target.value)}
                    style={{ appearance: "auto", borderColor: "#7E22CE" }}
                  >
                    <option value="">Selecione o Téc. de Segurança...</option>
                    {INSPECTORES_ENEL.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 8 }}>
                    {INSPECTORES_LIGHT.map((opt) => {
                      const active = form.quemInspecionou === opt;
                      return (
                        <div
                          key={opt}
                          className="radio-pill"
                          onClick={() => update("quemInspecionou", opt)}
                          style={{
                            background: active ? "#1F6B3A" : "#1C2126",
                            borderColor: active ? "#2F9E52" : "#2E3540",
                            color: active ? "#EAF4EE" : "#8A93A0",
                            fontWeight: 700,
                          }}
                        >
                          {opt}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>

              {/* Nome do Encarregado */}
              <div>
                <label className="field-label">Nome do Encarregado</label>
                <input
                  type="text"
                  className="field-input"
                  placeholder="Ex: João Silva"
                  value={form.encarregado}
                  onChange={(e) => update("encarregado", e.target.value)}
                />
              </div>

              {/* Eletricista Líder */}
              <div>
                <label className="field-label">Eletricista Líder</label>
                <input
                  type="text"
                  className="field-input"
                  placeholder="Ex: Carlos Pereira"
                  value={form.nomeEletricistaLider}
                  onChange={(e) => update("nomeEletricistaLider", e.target.value)}
                />
              </div>

              {/* Eletricista */}
              <div>
                <label className="field-label">Eletricista</label>
                <input
                  type="text"
                  className="field-input"
                  placeholder="Ex: Pedro Almeida"
                  value={form.nomeEletricista}
                  onChange={(e) => update("nomeEletricista", e.target.value)}
                />
              </div>
            </div>
          </div>

          {/* SEÇÃO 3: CONFORMIDADE & FOTOS */}
          <div className="form-section-card">
            <div className="form-section-header">
              <div className="form-section-title">
                <ShieldCheck size={16} /> 3. Avaliação & Registro Fotográfico
              </div>
            </div>

            <div className="grid-2" style={{ alignItems: "start" }}>
              {/* Conformidade */}
              <div>
                <label className="field-label">Status de Conformidade *</label>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  <div
                    className="radio-pill"
                    onClick={() => setForm((f) => ({ ...f, conformidade: "Conforme", descNaoConformidade: "" }))}
                    style={{
                      background: form.conformidade === "Conforme" ? "#1F6B3A" : "#1C2126",
                      borderColor: form.conformidade === "Conforme" ? "#2F9E52" : "#2E3540",
                      color: form.conformidade === "Conforme" ? "#EAF4EE" : "#8A93A0",
                      fontWeight: 700,
                      gap: 8,
                      minHeight: 48,
                    }}
                  >
                    <CheckCircle2 size={18} /> Conforme
                  </div>
                  <div
                    className="radio-pill"
                    onClick={() => update("conformidade", "Não conforme")}
                    style={{
                      background: form.conformidade === "Não conforme" ? "#7A2626" : "#1C2126",
                      borderColor: form.conformidade === "Não conforme" ? "#D64545" : "#2E3540",
                      color: form.conformidade === "Não conforme" ? "#FBEAEA" : "#8A93A0",
                      fontWeight: 700,
                      gap: 8,
                      minHeight: 48,
                    }}
                  >
                    <XCircle size={18} /> Não conforme
                  </div>
                </div>

                {form.conformidade === "Não conforme" && (
                  <div style={{ marginTop: 12 }}>
                    <label className="field-label" style={{ color: "#E8930C" }}>
                      <AlertTriangle size={13} />
                      Descreva o desvio / não conformidade *
                    </label>
                    <textarea
                      className="field-input"
                      style={{ resize: "vertical", minHeight: 75, fontFamily: "inherit", borderColor: "#E8930C" }}
                      placeholder="Descreva detalhadamente o desvio identificado na atividade..."
                      value={form.descNaoConformidade}
                      onChange={(e) => update("descNaoConformidade", e.target.value)}
                    />
                  </div>
                )}
              </div>

              {/* Registro de Foto */}
              <div>
                <label className="field-label">
                  <Camera size={13} /> Registro Fotográfico
                </label>
                <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 8 }}>
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
                          fontWeight: 700,
                          justifyContent: "flex-start",
                          padding: "10px 14px",
                          gap: 10,
                        }}
                      >
                        <div style={{ width: 10, height: 10, borderRadius: "50%", background: active ? s.dot : "#4A5462" }} />
                        {opt}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>

            {/* Observações */}
            <div style={{ marginTop: 14 }}>
              <label className="field-label">Observações Adicionais (Opcional)</label>
              <input
                type="text"
                className="field-input"
                placeholder="Qualquer observação pertinente sobre a inspeção..."
                value={form.observacao}
                onChange={(e) => update("observacao", e.target.value)}
              />
            </div>
          </div>

          {/* Submit Action */}
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "center", marginTop: 20 }}>
            <button className="btn-primary" disabled={!requiredOk} onClick={handleSubmit}>
              <Plus size={18} /> {editingId ? "Salvar Alterações" : "Adicionar Registro de OS"}
            </button>
            {!requiredOk && (
              <span style={{ fontSize: 12, color: "#8A93A0" }}>
                {form.conformidade === "Não conforme" && !form.descNaoConformidade.trim()
                  ? "⚠️ Preencha a descrição da não conformidade para salvar."
                  : "⚠️ Selecione a data da inspeção para salvar."}
              </span>
            )}
          </div>
        </div>

        {/* Table / Mobile Cards */}
        <div className="table-container">
          <div className="table-header-bar">
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", flex: 1 }}>
              <input
                type="text"
                placeholder="Buscar por OS, encarregado, regional, processo..."
                className="field-input"
                style={{ maxWidth: 360 }}
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

            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
              <div style={{ display: "flex", background: "#1C2126", padding: 3, borderRadius: 6, border: "1px solid #2E3540" }}>
                <button
                  type="button"
                  onClick={() => setMobileTableView("cards")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 11, fontWeight: 700,
                    background: mobileTableView === "cards" ? "#E8930C" : "transparent",
                    color: mobileTableView === "cards" ? "#14181C" : "#8A93A0",
                  }}
                >
                  Cards
                </button>
                <button
                  type="button"
                  onClick={() => setMobileTableView("table")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 11, fontWeight: 700,
                    background: mobileTableView === "table" ? "#E8930C" : "transparent",
                    color: mobileTableView === "table" ? "#14181C" : "#8A93A0",
                  }}
                >
                  Tabela
                </button>
              </div>
              <button className="btn-ghost" onClick={exportXlsx}>
                <Download size={14} /> XLS
              </button>
              <button className="btn-ghost" onClick={exportPdf}>
                <Download size={14} /> PDF
              </button>
            </div>
          </div>

          {/* Seletor Mobile Cards vs Tabela */}
          {mobileTableView === "cards" ? (
            <div style={{ padding: "16px 16px 4px 16px" }}>
              {pageRows.length === 0 ? (
                <div style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                  {rows.length === 0 ? "Nenhum registro ainda. Preencha o formulário acima para começar." : "Nenhum resultado para esse filtro."}
                </div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
                  {pageRows.map((r) => {
                    const fs = FOTO_STYLES[r.registroFoto] || FOTO_STYLES["Enviado"];
                    const cto = getContrato(r);
                    const isConforme = r.conformidade === "Conforme";

                    return (
                      <div key={r.id} className="record-card-mobile">
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                          <span style={{
                            fontSize: 11, fontWeight: 800, padding: "3px 8px", borderRadius: 4,
                            background: cto === "ENEL" ? "#2B1A4A" : "#1B2F3E",
                            color: cto === "ENEL" ? "#C084FC" : "#38BDF8",
                            border: `1px solid ${cto === "ENEL" ? "#7E22CE" : "#0284C7"}`
                          }}>
                            {cto}
                          </span>
                          <span style={{ fontSize: 12, color: "#8A93A0" }}>{formatDateForExport(r.data)}</span>
                        </div>

                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                          <span style={{ fontSize: 16, fontWeight: 800, color: "#E8930C" }}>
                            OS: {r.os || "S/N"}
                          </span>
                          <span style={{
                            fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 12,
                            background: isConforme ? "#1F6B3A" : "#7A2626",
                            color: isConforme ? "#EAF4EE" : "#FBEAEA",
                          }}>
                            {r.conformidade}
                          </span>
                        </div>

                        <div style={{ fontSize: 12, display: "grid", gap: 4, color: "#9AA4B2" }}>
                          <div><strong>Encarregado:</strong> {r.encarregado || "—"}</div>
                          <div><strong>Regional / Tipo:</strong> {r.regional || "—"} · {r.tipoInspecao || "—"}</div>
                          <div><strong>Inspetor:</strong> <span style={{ color: "#38BDF8" }}>{r.quemInspecionou || "—"}</span></div>
                          {r.processo && <div><strong>Processo:</strong> {r.processo}</div>}
                          {r.descNaoConformidade && (
                            <div style={{ color: "#FCA5A5", background: "#3A1B1B", padding: "6px 8px", borderRadius: 4, marginTop: 4 }}>
                              ⚠️ <strong>Desvio:</strong> {r.descNaoConformidade}
                            </div>
                          )}
                        </div>

                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderTop: "1px solid #28303B", paddingTop: 10, marginTop: 4 }}>
                          <span style={{ fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 12, background: fs.bg, color: fs.fg }}>
                            {r.registroFoto}
                          </span>
                          <div style={{ display: "flex", gap: 8 }}>
                            <button
                              onClick={() => handleEdit(r)}
                              style={{ background: "#252D37", border: "1px solid #3A4452", color: "#E8EBEE", borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 700, display: "flex", alignItems: "center", gap: 4 }}
                            >
                              <Edit3 size={13} /> Editar
                            </button>
                            <button
                              onClick={() => handleDelete(r.id)}
                              style={{ background: "#2D1A1A", border: "1px solid #5A2A2A", color: "#D64545", borderRadius: 6, padding: "6px 10px", cursor: "pointer" }}
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table>
                <thead>
                  <tr>
                    <th>Data</th>
                    <th>Contrato</th>
                    <th>Encarregado</th>
                    <th>OS</th>
                    <th>Tipo de Serviço</th>
                    <th>Regional</th>
                    <th>Status</th>
                    <th>Processo</th>
                    <th>Observação</th>
                    <th>Líder</th>
                    <th>Eletricista</th>
                    <th>Quem inspecionou</th>
                    <th>Tipo Inspeção</th>
                    <th>Arquivos Regras</th>
                    <th>Regras de Ouro</th>
                    <th>Foto</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {pageRows.length === 0 && (
                    <tr>
                      <td colSpan={17} style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                        {rows.length === 0 ? "Nenhum registro ainda. Preencha o formulário acima para começar." : "Nenhum resultado para esse filtro."}
                      </td>
                    </tr>
                  )}
                  {pageRows.map((r) => {
                    const fs = FOTO_STYLES[r.registroFoto] || FOTO_STYLES["Enviado"];
                    const cto = getContrato(r);
                    return (
                      <tr key={r.id}>
                        <td>{r.data}</td>
                        <td>
                          <span style={{
                            fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 4,
                            background: cto === "ENEL" ? "#2B1A4A" : "#1B2F3E",
                            color: cto === "ENEL" ? "#C084FC" : "#38BDF8",
                            border: `1px solid ${cto === "ENEL" ? "#7E22CE" : "#0284C7"}`
                          }}>
                            {cto}
                          </span>
                        </td>
                        <td>{r.encarregado || "—"}</td>
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
                        <td style={{ color: "#38BDF8", fontWeight: 600 }}>{r.quemInspecionou || "—"}</td>
                        <td>{r.tipoInspecao || "—"}</td>
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
          )}

          <div style={{ padding: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', borderTop: '1px solid #2E3540' }}>
            <div style={{ color: '#8A93A0', fontSize: 12 }}>
              Mostrando {pageRows.length} de {sortedFiltered.length} registros
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
              <button
                className="btn-ghost"
                disabled={currentPage === 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                style={{ opacity: currentPage === 1 ? 0.5 : 1 }}
              >
                Anterior
              </button>
              <span style={{ color: '#E8EBEE', fontSize: 12 }}>
                Página {currentPage} de {totalPages}
              </span>
              <button
                className="btn-ghost"
                disabled={currentPage === totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                style={{ opacity: currentPage === totalPages ? 0.5 : 1 }}
              >
                Próxima
              </button>
            </div>
          </div>
        </div>
      </>
      )}

      {/* ABA MONITORIA */}
      {view === "monitoria" && (
        <>
          <div className="card-box">
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 20, flexWrap: "wrap", gap: 10 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 8, height: 24, background: "#E8930C", borderRadius: 4 }}></div>
                <div className="disp" style={{ fontSize: 20, fontWeight: 800, letterSpacing: "0.02em" }}>
                  {editingId ? "EDITAR MONITORIA DE FILMAGEM" : "NOVA MONITORIA DE FILMAGEM"}
                </div>
              </div>
              {editingId && (
                <button className="btn-ghost" onClick={cancelEdit} style={{ height: 36, padding: "0 14px", fontSize: 12 }}>
                  Cancelar edição
                </button>
              )}
            </div>

            {/* SEÇÃO 1: CONTRATO & REGIONAL DA MONITORIA */}
            <div className="form-section-card">
              <div className="form-section-header">
                <div className="form-section-title">
                  <Building2 size={16} /> 1. Contrato, Regional & Veículo
                </div>
              </div>

              {/* Contrato Toggle */}
              <div style={{ marginBottom: 16 }}>
                <label className="field-label">Selecione o Contrato *</label>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  {['LIGHT', 'ENEL'].map((opt) => {
                    const active = (formMonitoria.contrato || "LIGHT") === opt;
                    const isLight = opt === 'LIGHT';
                    return (
                      <div
                        key={opt}
                        className="radio-pill"
                        onClick={() => handleContratoChangeMonitoria(opt)}
                        style={{
                          background: active 
                            ? (isLight ? "rgba(56, 189, 248, 0.18)" : "rgba(192, 132, 252, 0.18)") 
                            : "#1C2126",
                          borderColor: active 
                            ? (isLight ? "#38BDF8" : "#C084FC") 
                            : "#2E3540",
                          color: active 
                            ? (isLight ? "#38BDF8" : "#C084FC") 
                            : "#8A93A0",
                          fontSize: 14,
                          fontWeight: 800,
                          gap: 8,
                          boxShadow: active ? `0 0 12px ${isLight ? 'rgba(56, 189, 248, 0.25)' : 'rgba(192, 132, 252, 0.25)'}` : "none",
                        }}
                      >
                        {active && <Check size={16} />}
                        {opt}
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Regional Chips */}
              <div style={{ marginBottom: 16 }}>
                <label className="field-label">
                  <MapPin size={13} />
                  Regional ({(formMonitoria.contrato || "LIGHT")})
                </label>
                <div style={{ display: "grid", gridTemplateColumns: `repeat(${(REGIONAIS_POR_CONTRATO[formMonitoria.contrato || "LIGHT"] || []).length || 2}, 1fr)`, gap: 8 }}>
                  {(REGIONAIS_POR_CONTRATO[formMonitoria.contrato || "LIGHT"] || []).map((opt) => {
                    const active = formMonitoria.regional === opt;
                    return (
                      <div
                        key={opt}
                        className="radio-pill"
                        onClick={() => updateMonitoria("regional", opt)}
                        style={{
                          background: active ? "#1F6B3A" : "#1C2126",
                          borderColor: active ? "#2F9E52" : "#2E3540",
                          color: active ? "#EAF4EE" : "#8A93A0",
                          fontSize: 13,
                          fontWeight: 700,
                          padding: "8px 6px",
                        }}
                      >
                        {opt}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="grid-2" style={{ gridTemplateColumns: "1fr 1fr", margin: 0 }}>
                {/* Data Filmagem */}
                <div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                    <label className="field-label" style={{ margin: 0 }}>Data da Filmagem *</label>
                    <button type="button" onClick={setFormMonitoriaToday} className="btn-quick-today">
                      <Calendar size={12} /> Hoje
                    </button>
                  </div>
                  <input
                    type="date"
                    className="field-input"
                    value={formMonitoria.dataFilmagem}
                    onChange={(e) => updateMonitoria("dataFilmagem", e.target.value)}
                  />
                </div>

                {/* Placa do Veículo */}
                <div>
                  <label className="field-label">Placa do Veículo *</label>
                  <input
                    type="text"
                    className="field-input"
                    placeholder="Ex: ABC-1234 / ABC1D23"
                    value={formMonitoria.placaVeiculo}
                    onChange={(e) => updateMonitoria("placaVeiculo", e.target.value.toUpperCase())}
                    style={{ fontWeight: 700, color: "#E8930C" }}
                  />
                </div>
              </div>
            </div>

            {/* SEÇÃO 2: FISCAL, SETOR & SUPERVISÃO */}
            <div className="form-section-card">
              <div className="form-section-header">
                <div className="form-section-title">
                  <Users size={16} /> 2. Fiscalização, Setor & Equipe
                </div>
              </div>

              <div className="grid-3">
                <div>
                  <label className="field-label">Fiscal Responsável</label>
                  <select
                    className="field-input"
                    value={formMonitoria.fiscal}
                    onChange={(e) => updateMonitoria("fiscal", e.target.value)}
                    style={{ appearance: "auto" }}
                  >
                    <option value="">Selecione o fiscal...</option>
                    {FISCAIS_OPTIONS.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="field-label">Setor</label>
                  <select
                    className="field-input"
                    value={formMonitoria.setor}
                    onChange={(e) => updateMonitoria("setor", e.target.value)}
                    style={{ appearance: "auto" }}
                  >
                    <option value="">Selecione o setor...</option>
                    {SETORES_OPTIONS.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="field-label">Supervisor Responsável</label>
                  <select
                    className="field-input"
                    value={formMonitoria.supervisorResponsavel}
                    onChange={(e) => updateMonitoria("supervisorResponsavel", e.target.value)}
                    style={{ appearance: "auto" }}
                  >
                    <option value="">Selecione o supervisor...</option>
                    {SUPERVISORES_OPTIONS.map((name) => (
                      <option key={name} value={name}>{name}</option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Eletricistas e Motoristas Dinâmicos */}
              <div className="grid-2" style={{ gridTemplateColumns: "1fr 1fr", alignItems: "start", margin: 0 }}>
                {/* Eletricistas */}
                <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 16 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                    <label className="field-label" style={{ margin: 0, color: "#E8EBEE" }}>Eletricistas</label>
                    <button
                      type="button"
                      onClick={handleAddEletricista}
                      className="btn-ghost"
                      style={{ padding: "4px 10px", fontSize: 11, gap: 4 }}
                    >
                      <UserPlus size={13} /> Adicionar
                    </button>
                  </div>
                  <div style={{ display: "grid", gap: 10 }}>
                    {formMonitoria.eletricistas.map((ele, idx) => (
                      <div key={idx} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <input
                          type="text"
                          className="field-input"
                          placeholder={`Nome do Eletricista ${idx + 1}`}
                          value={ele}
                          onChange={(e) => handleUpdateEletricista(idx, e.target.value)}
                        />
                        {formMonitoria.eletricistas.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveEletricista(idx)}
                            style={{ background: "none", border: "1px solid #2E3540", color: "#D64545", borderRadius: 4, padding: "8px 10px", cursor: "pointer" }}
                            title="Remover eletricista"
                          >
                            <Trash2 size={14} />
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>

                {/* Motoristas */}
                <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 16 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                    <label className="field-label" style={{ margin: 0, color: "#E8EBEE" }}>Motoristas</label>
                    <button
                      type="button"
                      onClick={handleAddMotorista}
                      className="btn-ghost"
                      style={{ padding: "4px 10px", fontSize: 11, gap: 4 }}
                    >
                      <UserPlus size={13} /> Adicionar
                    </button>
                  </div>
                  <div style={{ display: "grid", gap: 10 }}>
                    {formMonitoria.motoristas.map((mot, idx) => (
                      <div key={idx} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <input
                          type="text"
                          className="field-input"
                          placeholder={`Nome do Motorista ${idx + 1}`}
                          value={mot}
                          onChange={(e) => handleUpdateMotorista(idx, e.target.value)}
                        />
                        {formMonitoria.motoristas.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveMotorista(idx)}
                            style={{ background: "none", border: "1px solid #2E3540", color: "#D64545", borderRadius: 4, padding: "8px 10px", cursor: "pointer" }}
                            title="Remover motorista"
                          >
                            <Trash2 size={14} />
                          </button>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>

            {/* SEÇÃO 3: AVALIAÇÃO DE CONFORMIDADE DA MONITORIA */}
            <div className="form-section-card">
              <div className="form-section-header">
                <div className="form-section-title">
                  <ShieldCheck size={16} /> 3. Avaliação de Conformidade & Tratativas
                </div>
              </div>

              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
                <label className="field-label" style={{ margin: 0, fontSize: 12, color: "#E8EBEE" }}>Houve Não Conformidades na Filmagem?</label>
                <div style={{ display: "flex", gap: 8 }}>
                  <div
                    className="radio-pill"
                    onClick={() => setFormMonitoria((f) => ({ ...f, houveNaoConformidade: "Não", descNaoConformidade: "", agenteAgressor: "", tratativas: "" }))}
                    style={{
                      background: formMonitoria.houveNaoConformidade === "Não" ? "#1F6B3A" : "#14181C",
                      borderColor: formMonitoria.houveNaoConformidade === "Não" ? "#2F9E52" : "#2E3540",
                      color: formMonitoria.houveNaoConformidade === "Não" ? "#EAF4EE" : "#8A93A0",
                      padding: "8px 18px",
                      gap: 6,
                    }}
                  >
                    <CheckCircle2 size={14} /> Não (Conforme)
                  </div>
                  <div
                    className="radio-pill"
                    onClick={() => updateMonitoria("houveNaoConformidade", "Sim")}
                    style={{
                      background: formMonitoria.houveNaoConformidade === "Sim" ? "#7A2626" : "#14181C",
                      borderColor: formMonitoria.houveNaoConformidade === "Sim" ? "#D64545" : "#2E3540",
                      color: formMonitoria.houveNaoConformidade === "Sim" ? "#FBEAEA" : "#8A93A0",
                      padding: "8px 18px",
                      gap: 6,
                    }}
                  >
                    <XCircle size={14} /> Sim (Não Conforme)
                  </div>
                </div>
              </div>

              {formMonitoria.houveNaoConformidade === "Sim" && (
                <div style={{ display: "grid", gap: 14, paddingTop: 10, borderTop: "1px solid #2E3540" }}>
                  <div>
                    <label className="field-label" style={{ color: "#E8930C" }}>
                      <AlertTriangle size={11} style={{ display: "inline", marginRight: 4, verticalAlign: -1 }} />
                      Descrição da Não Conformidade *
                    </label>
                    <textarea
                      className="field-input"
                      style={{ resize: "vertical", minHeight: 75, fontFamily: "inherit" }}
                      placeholder="Descreva detalhadamente a não conformidade observada na filmagem..."
                      value={formMonitoria.descNaoConformidade}
                      onChange={(e) => updateMonitoria("descNaoConformidade", e.target.value)}
                    />
                  </div>

                  <div className="grid-2" style={{ gridTemplateColumns: "1fr 1fr", margin: 0 }}>
                    <div>
                      <label className="field-label">Agente Agressor</label>
                      <select
                        className="field-input"
                        value={formMonitoria.agenteAgressor}
                        onChange={(e) => updateMonitoria("agenteAgressor", e.target.value)}
                        style={{ appearance: "auto" }}
                      >
                        <option value="">Selecione o agente agressor...</option>
                        {AGENTES_AGRESSORES_OPTIONS.map((opt) => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <label className="field-label">Tratativas</label>
                      <select
                        className="field-input"
                        value={formMonitoria.tratativas}
                        onChange={(e) => updateMonitoria("tratativas", e.target.value)}
                        style={{ appearance: "auto" }}
                      >
                        <option value="">Selecione a tratativa...</option>
                        {TRATATIVAS_OPTIONS.map((opt) => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              )}
            </div>

            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "center", marginTop: 20 }}>
              <button
                className="btn-primary"
                disabled={!formMonitoria.dataFilmagem || !formMonitoria.placaVeiculo || (formMonitoria.houveNaoConformidade === "Sim" && !formMonitoria.descNaoConformidade.trim())}
                onClick={handleSubmitMonitoria}
              >
                <Plus size={18} /> {editingId ? "Salvar Alterações" : "Registrar Monitoria"}
              </button>
              {(!formMonitoria.dataFilmagem || !formMonitoria.placaVeiculo) && (
                <span style={{ fontSize: 12, color: "#8A93A0" }}>
                  ⚠️ Preencha a data da filmagem e a placa do veículo para salvar.
                </span>
              )}
            </div>
          </div>

          {/* Tabela / Cards de Monitoria */}
          <div className="table-container">
            <div className="table-header-bar">
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", flex: 1 }}>
                <input
                  type="text"
                  placeholder="Filtrar por placa, fiscal, setor, supervisor, regional..."
                  className="field-input"
                  style={{ maxWidth: 360 }}
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

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                <div style={{ display: "flex", background: "#1C2126", padding: 3, borderRadius: 6, border: "1px solid #2E3540" }}>
                  <button
                    type="button"
                    onClick={() => setMobileMonitoriaView("cards")}
                    style={{
                      border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 11, fontWeight: 700,
                      background: mobileMonitoriaView === "cards" ? "#E8930C" : "transparent",
                      color: mobileMonitoriaView === "cards" ? "#14181C" : "#8A93A0",
                    }}
                  >
                    Cards
                  </button>
                  <button
                    type="button"
                    onClick={() => setMobileMonitoriaView("table")}
                    style={{
                      border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 11, fontWeight: 700,
                      background: mobileMonitoriaView === "table" ? "#E8930C" : "transparent",
                      color: mobileMonitoriaView === "table" ? "#14181C" : "#8A93A0",
                    }}
                  >
                    Tabela
                  </button>
                </div>
                <button className="btn-ghost" onClick={exportXlsx}>
                  <Download size={14} /> XLS
                </button>
                <button className="btn-ghost" onClick={exportPdf}>
                  <Download size={14} /> PDF
                </button>
              </div>
            </div>

            {mobileMonitoriaView === "cards" ? (
              <div style={{ padding: "16px 16px 4px 16px" }}>
                {pageRows.length === 0 ? (
                  <div style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                    Nenhum registro de monitoria encontrado.
                  </div>
                ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(300px, 1fr))", gap: 14 }}>
                    {pageRows.map((r) => {
                      const cto = getContrato(r);
                      const isNaoConforme = r.houveNaoConformidade === "Sim";
                      const eles = Array.isArray(r.eletricistas) ? r.eletricistas.filter(Boolean) : [];
                      const mots = Array.isArray(r.motoristas) ? r.motoristas.filter(Boolean) : [];

                      return (
                        <div key={r.id} className="record-card-mobile">
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                              <span style={{
                                fontSize: 11, fontWeight: 800, padding: "3px 8px", borderRadius: 4,
                                background: cto === "ENEL" ? "#2B1A4A" : "#1B2F3E",
                                color: cto === "ENEL" ? "#C084FC" : "#38BDF8",
                                border: `1px solid ${cto === "ENEL" ? "#7E22CE" : "#0284C7"}`
                              }}>
                                {cto}
                              </span>
                              {r.regional && (
                                <span style={{ fontSize: 11, fontWeight: 700, padding: "3px 8px", borderRadius: 4, background: "#1C2126", color: "#9AA4B2", border: "1px solid #2E3540" }}>
                                  {r.regional}
                                </span>
                              )}
                            </div>
                            <span style={{ fontSize: 12, color: "#8A93A0" }}>{formatDateForExport(r.dataFilmagem || r.data)}</span>
                          </div>

                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline" }}>
                            <span style={{ fontSize: 16, fontWeight: 800, color: "#E8930C" }}>
                              {r.placaVeiculo || "PLACA S/N"}
                            </span>
                            <span style={{
                              fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 12,
                              background: !isNaoConforme ? "#1F6B3A" : "#7A2626",
                              color: !isNaoConforme ? "#EAF4EE" : "#FBEAEA",
                            }}>
                              {isNaoConforme ? "Não Conforme" : "Conforme"}
                            </span>
                          </div>

                          <div style={{ fontSize: 12, display: "grid", gap: 4, color: "#9AA4B2" }}>
                            <div><strong>Fiscal:</strong> <span style={{ color: "#38BDF8" }}>{r.fiscal || "—"}</span></div>
                            <div><strong>Setor / Supervisor:</strong> {r.setor || "—"} · {r.supervisorResponsavel || "—"}</div>
                            {eles.length > 0 && <div><strong>Eletricistas:</strong> {eles.join(", ")}</div>}
                            {mots.length > 0 && <div><strong>Motoristas:</strong> {mots.join(", ")}</div>}
                            {r.descNaoConformidade && (
                              <div style={{ color: "#FCA5A5", background: "#3A1B1B", padding: "6px 8px", borderRadius: 4, marginTop: 4 }}>
                                ⚠️ <strong>Não conformidade:</strong> {r.descNaoConformidade}
                              </div>
                            )}
                            {r.agenteAgressor && (
                              <div><strong>Agente:</strong> <span style={{ color: "#E8930C" }}>{r.agenteAgressor}</span></div>
                            )}
                            {r.tratativas && <div><strong>Tratativa:</strong> {r.tratativas}</div>}
                          </div>

                          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, borderTop: "1px solid #28303B", paddingTop: 10, marginTop: 4 }}>
                            <button
                              onClick={() => handleEdit(r)}
                              style={{ background: "#252D37", border: "1px solid #3A4452", color: "#E8EBEE", borderRadius: 6, padding: "6px 12px", fontSize: 12, cursor: "pointer", fontWeight: 700, display: "flex", alignItems: "center", gap: 4 }}
                            >
                              <Edit3 size={13} /> Editar
                            </button>
                            <button
                              onClick={() => handleDelete(r.id)}
                              style={{ background: "#2D1A1A", border: "1px solid #5A2A2A", color: "#D64545", borderRadius: 6, padding: "6px 10px", cursor: "pointer" }}
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            ) : (
              <div style={{ overflowX: "auto" }}>
                <table>
                  <thead>
                    <tr>
                      <th>Data Filmagem</th>
                      <th>Contrato</th>
                      <th>Regional</th>
                      <th>Placa</th>
                      <th>Fiscal</th>
                      <th>Setor</th>
                      <th>Supervisor</th>
                      <th>Eletricistas</th>
                      <th>Motoristas</th>
                      <th>Conformidade</th>
                      <th>Não Conformidade</th>
                      <th>Agente Agressor</th>
                      <th>Tratativas</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {pageRows.length === 0 && (
                      <tr>
                        <td colSpan={14} style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                          Nenhum registro de monitoria encontrado.
                        </td>
                      </tr>
                    )}
                    {pageRows.map((r) => {
                      const cto = getContrato(r);
                      const isNaoConforme = r.houveNaoConformidade === "Sim";
                      const eles = Array.isArray(r.eletricistas) ? r.eletricistas.filter(Boolean) : [];
                      const mots = Array.isArray(r.motoristas) ? r.motoristas.filter(Boolean) : [];
                      return (
                        <tr key={r.id}>
                          <td>{formatDateForExport(r.dataFilmagem || r.data)}</td>
                          <td>
                            <span style={{
                              fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 4,
                              background: cto === "ENEL" ? "#2B1A4A" : "#1B2F3E",
                              color: cto === "ENEL" ? "#C084FC" : "#38BDF8",
                              border: `1px solid ${cto === "ENEL" ? "#7E22CE" : "#0284C7"}`
                            }}>
                              {cto}
                            </span>
                          </td>
                          <td>{r.regional || "—"}</td>
                          <td style={{ color: "#E8930C", fontWeight: 700 }}>{r.placaVeiculo}</td>
                          <td>{r.fiscal || "—"}</td>
                          <td>
                            <span style={{ fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 4, background: "#1B2F3E", color: "#38BDF8" }}>
                              {r.setor || "—"}
                            </span>
                          </td>
                          <td>{r.supervisorResponsavel || "—"}</td>
                          <td>
                            <div style={{ display: "flex", flexDirection: "column", gap: 2, fontSize: 12 }}>
                              {eles.length ? eles.map((name, i) => <span key={i}>• {name}</span>) : "—"}
                            </div>
                          </td>
                          <td>
                            <div style={{ display: "flex", flexDirection: "column", gap: 2, fontSize: 12 }}>
                              {mots.length ? mots.map((name, i) => <span key={i}>• {name}</span>) : "—"}
                            </div>
                          </td>
                          <td>
                            <span style={{
                              fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 20,
                              background: !isNaoConforme ? "#1F6B3A" : "#7A2626",
                              color: !isNaoConforme ? "#EAF4EE" : "#FBEAEA",
                            }}>
                              {isNaoConforme ? "Não Conforme" : "Conforme"}
                            </span>
                          </td>
                          <td style={{ maxWidth: 220, fontSize: 12, color: isNaoConforme ? "#FBEAEA" : "#6B7580" }}>
                            {r.descNaoConformidade || "—"}
                          </td>
                          <td>
                            {r.agenteAgressor ? (
                              <span style={{ fontSize: 11, fontWeight: 600, padding: "2px 6px", borderRadius: 4, background: "#2A2016", color: "#E8930C", border: "1px solid #784807" }}>
                                {r.agenteAgressor}
                              </span>
                            ) : "—"}
                          </td>
                          <td style={{ fontSize: 12 }}>{r.tratativas || "—"}</td>
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
            )}

            <div style={{ padding: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', borderTop: '1px solid #2E3540' }}>
              <div style={{ color: '#8A93A0', fontSize: 12 }}>
                Mostrando {pageRows.length} de {sortedFiltered.length} registros
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <button
                  className="btn-ghost"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  style={{ opacity: currentPage === 1 ? 0.5 : 1 }}
                >
                  Anterior
                </button>
                <span style={{ color: '#E8EBEE', fontSize: 12 }}>
                  Página {currentPage} de {totalPages}
                </span>
                <button
                  className="btn-ghost"
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  style={{ opacity: currentPage === totalPages ? 0.5 : 1 }}
                >
                  Próxima
                </button>
              </div>
            </div>
          </div>
        </>
      )}

      {/* ABA DASHBOARD */}
      {view === "dashboard" && (
        <DashboardView
          dash={dash}
          dashMonitoria={dashMonitoria}
          rowsCount={dashboardRows.length}
          monitoriaCount={monitoriaDashboardRows.length}
          monthOptions={monthOptions}
          dashboardMonth={dashboardMonth}
          setDashboardMonth={setDashboardMonth}
          dashboardContrato={dashboardContrato}
          setDashboardContrato={setDashboardContrato}
          dashboardSetor={dashboardSetor}
          setDashboardSetor={setDashboardSetor}
          contratoStats={contratoStats}
          contratoStatsMonitoria={contratoStatsMonitoria}
          dashboardTab={dashboardTab}
          setDashboardTab={setDashboardTab}
        />
      )}

      {/* ABA NÃO ENVIO */}
      {view === "naoEnvio" && (
        <div className="card-box">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18, flexWrap: "wrap", gap: 10 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 8, height: 24, background: "#E8930C", borderRadius: 4 }}></div>
              <div className="disp" style={{ fontSize: 20, fontWeight: 800, letterSpacing: "0.02em" }}>
                {editingId ? "EDITAR NÃO ENVIO DE RO" : "REGISTRO DE NÃO ENVIO DE RO"}
              </div>
            </div>
            {editingId && (
              <button className="btn-ghost" onClick={cancelEdit} style={{ height: 36, padding: "0 14px", fontSize: 12 }}>
                Cancelar edição
              </button>
            )}
          </div>

          <div className="grid-4">
            <div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
                <label className="field-label" style={{ margin: 0 }}>Data</label>
                <button type="button" onClick={() => update("data", getTodayDate())} className="btn-quick-today">
                  <Calendar size={12} /> Hoje
                </button>
              </div>
              <input
                type="date"
                className="field-input"
                value={form.data}
                onChange={(e) => setForm((f) => ({ ...f, data: e.target.value }))}
              />
            </div>
            <div>
              <label className="field-label">Nome do Encarregado</label>
              <input
                type="text"
                className="field-input"
                placeholder="Ex: João Silva"
                value={form.encarregado}
                onChange={(e) => setForm((f) => ({ ...f, encarregado: e.target.value }))}
              />
            </div>
            <div>
              <label className="field-label">Filial</label>
              <input
                type="text"
                className="field-input"
                placeholder="Ex: Baixada, Oeste, Cantagalo..."
                value={form.filial}
                onChange={(e) => setForm((f) => ({ ...f, filial: e.target.value }))}
              />
            </div>
            <div>
              <label className="field-label">Número da OS</label>
              <input
                type="text"
                className="field-input"
                placeholder="Nº da OS"
                value={form.os}
                onChange={(e) => setForm((f) => ({ ...f, os: e.target.value }))}
                style={{ fontWeight: 600, color: "#E8930C" }}
              />
            </div>
          </div>

          <div style={{ marginBottom: 20 }}>
            <label className="field-label">Descrição do Motivo de Não Envio</label>
            <textarea
              className="field-input"
              style={{ minHeight: 90, resize: "vertical", fontFamily: "inherit" }}
              placeholder="Informe o motivo ou justificativa do não envio da RO..."
              value={form.naoEnvio}
              onChange={(e) => setForm((f) => ({ ...f, naoEnvio: e.target.value }))}
            />
          </div>

          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
            <button
              className="btn-primary"
              onClick={handleSubmit}
            >
              <Plus size={18} /> {editingId ? "Salvar Alterações" : "Registrar Não Envio"}
            </button>
            <button className="btn-ghost" onClick={() => { setForm(EMPTY_NAO_ENVIO); setEditingId(null); }}>
              Limpar Campos
            </button>
          </div>

          <div className="grid-3" style={{ marginTop: 24 }}>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 10, padding: 18 }}>
              <div style={{ fontSize: 12, color: "#8A93A0", marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.06em" }}>Total de Não Envios</div>
              <div className="disp" style={{ fontSize: 32, fontWeight: 800, color: "#E8930C" }}>{naoEnvioSummary.total}</div>
            </div>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 10, padding: 18 }}>
              <div style={{ fontSize: 12, color: "#8A93A0", marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.06em" }}>Encarregados Hoje</div>
              <div className="disp" style={{ fontSize: 32, fontWeight: 800, color: "#2F9E52" }}>{naoEnvioSummary.today}</div>
              <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 6 }}>{naoEnvioSummary.todayLabel}</div>
            </div>
            <div style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 10, padding: 18 }}>
              <div style={{ fontSize: 12, color: "#8A93A0", marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.06em" }}>Encarregados Esta Semana</div>
              <div className="disp" style={{ fontSize: 32, fontWeight: 800, color: "#4D8FFF" }}>{naoEnvioSummary.week}</div>
              <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 6 }}>{naoEnvioSummary.currentWeek}</div>
            </div>
          </div>

          <div className="table-container" style={{ marginTop: 24 }}>
            <div className="table-header-bar">
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", flex: 1 }}>
                <input
                  type="text"
                  placeholder="Filtrar não envio por OS, encarregado, filial..."
                  className="field-input"
                  style={{ maxWidth: 360 }}
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
                <button className="btn-ghost" onClick={exportXlsx}>
                  <Download size={14} /> XLS
                </button>
                <button className="btn-ghost" onClick={exportPdf}>
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
                    <th>Filial</th>
                    <th>OS</th>
                    <th>Motivo do Não Envio</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {pageRows.length === 0 && (
                    <tr>
                      <td colSpan={6} style={{ textAlign: "center", color: "#6B7580", padding: "32px 12px" }}>
                        Nenhum registro de não envio encontrado.
                      </td>
                    </tr>
                  )}
                  {pageRows.map((r) => (
                    <tr key={r.id}>
                      <td>{formatDateForExport(r.data)}</td>
                      <td>{r.encarregado || "—"}</td>
                      <td>{r.filial || "—"}</td>
                      <td style={{ color: "#E8930C", fontWeight: 700 }}>{r.os || "—"}</td>
                      <td style={{ maxWidth: 340 }}>{r.naoEnvio || "—"}</td>
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
                  ))}
                </tbody>
              </table>
            </div>
            <div style={{ padding: 16, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, flexWrap: 'wrap', borderTop: '1px solid #2E3540' }}>
              <div style={{ color: '#8A93A0', fontSize: 12 }}>
                Mostrando {pageRows.length} de {sortedFiltered.length} registros
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <button
                  className="btn-ghost"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                  style={{ opacity: currentPage === 1 ? 0.5 : 1 }}
                >
                  Anterior
                </button>
                <span style={{ color: '#E8EBEE', fontSize: 12 }}>
                  Página {currentPage} de {totalPages}
                </span>
                <button
                  className="btn-ghost"
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                  style={{ opacity: currentPage === totalPages ? 0.5 : 1 }}
                >
                  Próxima
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ABA REGRAS DE OURO */}
      {view === "regras" && (
        <div className="card-box">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 18 }}>
            <div className="disp" style={{ fontSize: 17, fontWeight: 700, letterSpacing: "0.01em" }}>
              REGRAS DE OURO / NR10
            </div>
          </div>
          <div className="grid-3">
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
          <div className="grid-3" style={{ marginBottom: 20 }}>
            {Object.entries(RULES_DE_OURO).map(([key, label]) => (
              <div key={key} style={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 8, padding: 18 }}>
                <div style={{ fontSize: 13, color: "#E8EBEE", marginBottom: 10, fontWeight: 700 }}>{label}</div>
                <div style={{ fontSize: 12, color: "#8A93A0", marginBottom: 14 }}>Registros com essa regra:</div>
                <div className="disp" style={{ fontSize: 28, fontWeight: 800, color: "#E8930C" }}>{ruleCounts[key] || 0}</div>
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

function DashboardView({
  dash,
  dashMonitoria,
  rowsCount,
  monitoriaCount,
  monthOptions,
  dashboardMonth,
  setDashboardMonth,
  dashboardContrato,
  setDashboardContrato,
  dashboardSetor,
  setDashboardSetor,
  contratoStats,
  contratoStatsMonitoria,
  dashboardTab,
  setDashboardTab,
}) {
  const cardStyle = { background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 20 };
  const titleStyle = { fontSize: 12, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", color: "#8A93A0", marginBottom: 16 };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {/* Sub-tabs no Dashboard */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16, borderBottom: "1px solid #2E3540", paddingBottom: 16 }}>
        <div style={{ display: "flex", gap: 8, background: "#1C2126", padding: 4, borderRadius: 6, border: "1px solid #2E3540" }}>
          <button
            onClick={() => setDashboardTab("os")}
            style={{
              display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
              padding: "8px 16px", borderRadius: 4, fontSize: 13, fontWeight: 700,
              background: dashboardTab === "os" ? "#E8930C" : "transparent",
              color: dashboardTab === "os" ? "#14181C" : "#8A93A0",
              transition: "all 0.15s",
            }}
          >
            <ClipboardList size={15} /> Dashboard OS ({rowsCount})
          </button>
          <button
            onClick={() => setDashboardTab("monitoria")}
            style={{
              display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer",
              padding: "8px 16px", borderRadius: 4, fontSize: 13, fontWeight: 700,
              background: dashboardTab === "monitoria" ? "#E8930C" : "transparent",
              color: dashboardTab === "monitoria" ? "#14181C" : "#8A93A0",
              transition: "all 0.15s",
            }}
          >
            <Film size={15} /> Dashboard Monitoria ({monitoriaCount})
          </button>
        </div>

        {/* Filtro de mês compartilhado */}
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 12, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.08em" }}>Filtrar mês</span>
          <select
            value={dashboardMonth}
            onChange={(e) => setDashboardMonth(e.target.value)}
            className="field-input"
            style={{ maxWidth: 180 }}
          >
            <option value="all">Todos os meses</option>
            {monthOptions.map((month) => (
              <option key={month} value={month}>{formatMonthLabel(month)}</option>
            ))}
          </select>
        </div>
      </div>

      {/* DASHBOARD DE REGISTROS DE OS */}
      {dashboardTab === "os" && (
        <>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 16, justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 12, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.08em" }}>Contrato</span>
              <div style={{ display: "flex", gap: 6, background: "#1C2126", padding: 3, borderRadius: 6, border: "1px solid #2E3540" }}>
                <button
                  onClick={() => setDashboardContrato("all")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "all" ? "#E8930C" : "transparent",
                    color: dashboardContrato === "all" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  Todos ({contratoStats.total})
                </button>
                <button
                  onClick={() => setDashboardContrato("LIGHT")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "LIGHT" ? "#38BDF8" : "transparent",
                    color: dashboardContrato === "LIGHT" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  LIGHT ({contratoStats.light})
                </button>
                <button
                  onClick={() => setDashboardContrato("ENEL")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "ENEL" ? "#C084FC" : "transparent",
                    color: dashboardContrato === "ENEL" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  ENEL ({contratoStats.enel})
                </button>
              </div>
            </div>

            <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
              <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>Regras de Ouro</div>
                <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#E8930C" }}>{dash.goldRuleCount}</div>
                <div style={{ fontSize: 11, color: "#8A93A0" }}>{dash.goldRuleRate}% dos registros</div>
              </div>
              {(dashboardContrato === "all" || dashboardContrato === "LIGHT") && (
                <>
                  <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>Proteção</div>
                    <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#2F9E52" }}>{dash.protecaoCount}</div>
                    <div style={{ fontSize: 11, color: "#8A93A0" }}>registros (LIGHT)</div>
                  </div>
                  <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>RDA</div>
                    <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#4D8FFF" }}>{dash.rdaCount}</div>
                    <div style={{ fontSize: 11, color: "#8A93A0" }}>registros (LIGHT)</div>
                  </div>
                </>
              )}
              {(dashboardContrato === "all" || dashboardContrato === "ENEL") && (
                <>
                  <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>Obras</div>
                    <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#E8930C" }}>{dash.obrasCount}</div>
                    <div style={{ fontSize: 11, color: "#8A93A0" }}>registros (ENEL)</div>
                  </div>
                  <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>Emergencial</div>
                    <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#F97316" }}>{dash.emergencialCount}</div>
                    <div style={{ fontSize: 11, color: "#8A93A0" }}>registros (ENEL)</div>
                  </div>
                  <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: 14, minWidth: 140 }}>
                    <div style={{ fontSize: 11, color: "#6B7580", marginBottom: 4, textTransform: "uppercase", letterSpacing: "0.08em" }}>Comercial</div>
                    <div className="disp" style={{ fontSize: 24, fontWeight: 800, color: "#10B981" }}>{dash.comercialCount}</div>
                    <div style={{ fontSize: 11, color: "#8A93A0" }}>registros (ENEL)</div>
                  </div>
                </>
              )}
            </div>
          </div>

          {rowsCount === 0 ? (
            <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: "60px 24px", textAlign: "center", color: "#6B7580" }}>
              Nenhum registro de OS encontrado para o filtro selecionado.
            </div>
          ) : (
            <>
              {/* KPI row */}
              <div className="grid-kpi">
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
              <div className="grid-charts-main">
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

              {/* Timeline + top encarregados */}
              <div className="grid-charts-sub">
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
            </>
          )}
        </>
      )}

      {/* DASHBOARD DE MONITORIA */}
      {dashboardTab === "monitoria" && (
        <>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 16, justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 12, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.08em" }}>Contrato</span>
              <div style={{ display: "flex", gap: 6, background: "#1C2126", padding: 3, borderRadius: 6, border: "1px solid #2E3540" }}>
                <button
                  onClick={() => setDashboardContrato("all")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "all" ? "#E8930C" : "transparent",
                    color: dashboardContrato === "all" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  Todos ({contratoStatsMonitoria.total})
                </button>
                <button
                  onClick={() => setDashboardContrato("LIGHT")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "LIGHT" ? "#38BDF8" : "transparent",
                    color: dashboardContrato === "LIGHT" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  LIGHT ({contratoStatsMonitoria.light})
                </button>
                <button
                  onClick={() => setDashboardContrato("ENEL")}
                  style={{
                    border: "none", cursor: "pointer", padding: "6px 12px", borderRadius: 4, fontSize: 12, fontWeight: 700,
                    background: dashboardContrato === "ENEL" ? "#C084FC" : "transparent",
                    color: dashboardContrato === "ENEL" ? "#14181C" : "#8A93A0",
                    transition: "all 0.15s",
                  }}
                >
                  ENEL ({contratoStatsMonitoria.enel})
                </button>
              </div>
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 12, color: "#6B7580", textTransform: "uppercase", letterSpacing: "0.08em" }}>Filtrar Setor</span>
              <select
                value={dashboardSetor}
                onChange={(e) => setDashboardSetor(e.target.value)}
                className="field-input"
                style={{ maxWidth: 200 }}
              >
                <option value="all">Todos os setores</option>
                {SETORES_OPTIONS.map((s) => (
                  <option key={s} value={s}>{s}</option>
                ))}
              </select>
            </div>
          </div>

          {monitoriaCount === 0 ? (
            <div style={{ background: "#171B20", border: "1px solid #2E3540", borderRadius: 8, padding: "60px 24px", textAlign: "center", color: "#6B7580" }}>
              Nenhum registro de monitoria encontrado para o filtro selecionado.
            </div>
          ) : (
            <>
              {/* KPI row Monitoria */}
              <div className="grid-kpi">
                <div style={cardStyle}>
                  <div style={titleStyle}>Taxa de conformidade</div>
                  <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: dashMonitoria.taxaConformidade >= 80 ? "#2F9E52" : dashMonitoria.taxaConformidade >= 50 ? "#E8930C" : "#D64545" }}>
                    {dashMonitoria.taxaConformidade}%
                  </div>
                  <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 4 }}>{dashMonitoria.conformeCount} filmagens conformes</div>
                </div>
                <div style={cardStyle}>
                  <div style={titleStyle}>Total de monitorias</div>
                  <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: "#E8EBEE" }}>{dashMonitoria.total}</div>
                  <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 4 }}>filmagens analisadas</div>
                </div>
                <div style={cardStyle}>
                  <div style={titleStyle}>Não conformidades</div>
                  <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: "#D64545" }}>{dashMonitoria.naoConformeCount}</div>
                  <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 4 }}>desvios detectados</div>
                </div>
                <div style={cardStyle}>
                  <div style={titleStyle}>Fiscais ativos</div>
                  <div className="disp" style={{ fontSize: 38, fontWeight: 800, color: "#38BDF8" }}>{dashMonitoria.fiscalData.length}</div>
                  <div style={{ fontSize: 11, color: "#8A93A0", marginTop: 4 }}>fiscais com registros</div>
                </div>
              </div>

              {/* Setor bar + Regional bar */}
              <div className="grid-charts-main">
                <div style={cardStyle}>
                  <div style={titleStyle}>Filmagens por Setor</div>
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={dashMonitoria.setorData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" vertical={false} />
                      <XAxis dataKey="name" stroke="#6B7580" fontSize={10} tickLine={false} axisLine={{ stroke: "#2E3540" }} />
                      <YAxis stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
                      <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} labelStyle={{ color: "#E8EBEE" }} />
                      <Legend wrapperStyle={{ fontSize: 12 }} />
                      <Bar dataKey="Conforme" stackId="a" fill={CHART_COLORS.conforme} radius={[0, 0, 0, 0]} />
                      <Bar dataKey="Não conforme" stackId="a" fill={CHART_COLORS.naoConforme} radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={cardStyle}>
                  <div style={titleStyle}>Filmagens por Regional</div>
                  {dashMonitoria.regionalData.length === 0 ? (
                    <div style={{ height: 260, display: "flex", alignItems: "center", justifyContent: "center", color: "#6B7580", fontSize: 12 }}>
                      Nenhuma regional informada no período.
                    </div>
                  ) : (
                    <ResponsiveContainer width="100%" height={260}>
                      <BarChart data={dashMonitoria.regionalData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" vertical={false} />
                        <XAxis dataKey="name" stroke="#6B7580" fontSize={10} tickLine={false} axisLine={{ stroke: "#2E3540" }} />
                        <YAxis stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
                        <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} labelStyle={{ color: "#E8EBEE" }} />
                        <Legend wrapperStyle={{ fontSize: 12 }} />
                        <Bar dataKey="Conforme" stackId="a" fill={CHART_COLORS.conforme} radius={[0, 0, 0, 0]} />
                        <Bar dataKey="Não conforme" stackId="a" fill={CHART_COLORS.naoConforme} radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>
              </div>

              {/* Agente Agressor + Tratativas + Top Fiscais */}
              <div className="grid-3">
                <div style={cardStyle}>
                  <div style={titleStyle}>Agentes Agressores</div>
                  {dashMonitoria.agenteData.length === 0 ? (
                    <div style={{ height: 240, display: "flex", alignItems: "center", justifyContent: "center", color: "#6B7580", fontSize: 12 }}>
                      Nenhum agente registrado.
                    </div>
                  ) : (
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={dashMonitoria.agenteData} layout="vertical" margin={{ left: 8 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" horizontal={false} />
                        <XAxis type="number" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
                        <YAxis dataKey="name" type="category" stroke="#6B7580" fontSize={10} tickLine={false} axisLine={{ stroke: "#2E3540" }} width={90} />
                        <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} />
                        <Bar dataKey="value" fill={CHART_COLORS.rose} radius={[0, 4, 4, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>

                <div style={cardStyle}>
                  <div style={titleStyle}>Tratativas Aplicadas</div>
                  {dashMonitoria.tratativaData.length === 0 ? (
                    <div style={{ height: 240, display: "flex", alignItems: "center", justifyContent: "center", color: "#6B7580", fontSize: 12 }}>
                      Nenhuma tratativa registrada.
                    </div>
                  ) : (
                    <ResponsiveContainer width="100%" height={240}>
                      <BarChart data={dashMonitoria.tratativaData}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" vertical={false} />
                        <XAxis dataKey="name" stroke="#6B7580" fontSize={10} tickLine={false} axisLine={{ stroke: "#2E3540" }} />
                        <YAxis stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
                        <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} labelStyle={{ color: "#E8EBEE" }} />
                        <Bar dataKey="value" fill={CHART_COLORS.accent} radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  )}
                </div>

                <div style={cardStyle}>
                  <div style={titleStyle}>Top Fiscais (Filmagens)</div>
                  <ResponsiveContainer width="100%" height={240}>
                    <BarChart data={dashMonitoria.fiscalData} layout="vertical" margin={{ left: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#2E3540" horizontal={false} />
                      <XAxis type="number" stroke="#6B7580" fontSize={11} tickLine={false} axisLine={{ stroke: "#2E3540" }} allowDecimals={false} />
                      <YAxis dataKey="name" type="category" stroke="#6B7580" fontSize={10} tickLine={false} axisLine={{ stroke: "#2E3540" }} width={90} />
                      <Tooltip contentStyle={{ background: "#1C2126", border: "1px solid #2E3540", borderRadius: 6, fontSize: 12 }} />
                      <Bar dataKey="value" fill={CHART_COLORS.purple} radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}
        </>
      )}
    </div>
  );
}
