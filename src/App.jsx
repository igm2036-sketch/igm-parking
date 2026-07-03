import React, { useState, useMemo, useEffect, useRef } from "react";
import { Car, Users, ClipboardList, Download, Check, X, Clock, ShieldCheck, Bell, Search, Lock, Trash2 } from "lucide-react";
import * as XLSX from "xlsx";
import { db } from "./firebase";
import {
  collection,
  addDoc,
  updateDoc,
  deleteDoc,
  doc,
  onSnapshot,
  query,
  orderBy,
} from "firebase/firestore";

const LOGO_SRC = "data:image/png;base64,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";

const TYPE_META = {
  visitor: { label: "커니 직원", icon: Car, color: "#F2A93B" },
  employee: { label: "IGM 직원", icon: Users, color: "#5B8DEF" },
};

const STATUS_META = {
  pending: { label: "대기", color: "#8A7B5C", bg: "#F4EEDD" },
  approved: { label: "승인", color: "#2F7D52", bg: "#E4F3EA" },
  rejected: { label: "거절", color: "#C0392B", bg: "#FBEAE7" },
};

function GateIcon({ status }) {
  // signature element: parking barrier gate that raises on approval
  const angle = status === "approved" ? -72 : 0;
  const barColor = status === "rejected" ? "#C0392B" : "#16233F";
  return (
    <svg width="46" height="40" viewBox="0 0 46 40" aria-hidden="true">
      <rect x="4" y="8" width="7" height="28" rx="1.5" fill="#16233F" />
      <rect x="35" y="8" width="7" height="28" rx="1.5" fill="#16233F" />
      <g style={{ transformOrigin: "7px 12px", transition: "transform 0.5s cubic-bezier(.4,1.6,.5,1)" }}
         transform={`rotate(${angle} 7 12)`}>
        <rect x="7" y="8.5" width="34" height="6" rx="3" fill={barColor} />
        <rect x="12" y="9.5" width="5" height="4" fill="#F2A93B" />
        <rect x="22" y="9.5" width="5" height="4" fill="#F2A93B" />
        <rect x="32" y="9.5" width="5" height="4" fill="#F2A93B" />
      </g>
    </svg>
  );
}

function FormField({ label, hint, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 13 }}>
      <span style={{ color: "#5B5646", fontWeight: 600, letterSpacing: "0.01em" }}>
        {label}
        {hint && (
          <span style={{ color: "#A79F8A", fontWeight: 500, marginLeft: 5, fontSize: 11.5 }}>
            {hint}
          </span>
        )}
      </span>
      {children}
    </label>
  );
}

const inputStyle = {
  border: "1.5px solid #E4DFD1",
  borderRadius: 8,
  padding: "10px 12px",
  fontSize: 14,
  fontFamily: "inherit",
  color: "#1C1C1E",
  background: "#FFFFFF",
  outline: "none",
  width: "100%",
  boxSizing: "border-box",
};

function formatPhone(raw) {
  const digits = raw.replace(/\D/g, "").slice(0, 11);
  if (digits.length < 4) return digits;
  if (digits.length < 7) return `${digits.slice(0, 3)}-${digits.slice(3)}`;
  if (digits.length < 11) return `${digits.slice(0, 3)}-${digits.slice(3, 6)}-${digits.slice(6)}`;
  return `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7, 11)}`;
}

export default function ParkingApp() {
  const [view, setView] = useState("form"); // 'form' | 'admin' | 'lookup'
  const [reqType, setReqType] = useState("visitor");
  const [requests, setRequests] = useState([]);
  const [filter, setFilter] = useState("all");
  const [toast, setToast] = useState(null);
  const [unseenCount, setUnseenCount] = useState(0);
  const [lookupQuery, setLookupQuery] = useState({ name: "", plate: "" });
  const [confirmation, setConfirmation] = useState(null);
  const [missingFields, setMissingFields] = useState(null);
  const [adminUnlocked, setAdminUnlocked] = useState(false);
  const [showPasswordPrompt, setShowPasswordPrompt] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [passwordInput, setPasswordInput] = useState("");
  const [passwordError, setPasswordError] = useState(false);
  const [loading, setLoading] = useState(true);
  const isFirstLoad = useRef(true);

  const ADMIN_PASSWORD = "2036";

  const todayStr = () => new Date().toISOString().slice(0, 10);

  const [form, setForm] = useState({
    name: "",
    plate: "",
    phone: "",
    date: todayStr(),
    dept: "",
    purpose: "",
  });

  // Firestore 실시간 구독: 누가 어디서 신청/승인해도 모든 기기에 즉시 반영돼요.
  useEffect(() => {
    const q = query(collection(db, "requests"), orderBy("submittedAt", "desc"));
    const unsubscribe = onSnapshot(
      q,
      (snapshot) => {
        const list = snapshot.docs.map((d) => ({ id: d.id, ...d.data() }));
        setRequests(list);
        if (!isFirstLoad.current) {
          const addedCount = snapshot.docChanges().filter((c) => c.type === "added").length;
          if (addedCount > 0) setUnseenCount((n) => n + addedCount);
        }
        isFirstLoad.current = false;
        setLoading(false);
      },
      (error) => {
        console.error("Firestore 구독 오류:", error);
        setLoading(false);
        showToast("데이터를 불러오지 못했어요. 잠시 후 새로고침해주세요.");
      }
    );
    return () => unsubscribe();
  }, []);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2400);
  };

  const resetForm = () =>
    setForm({ name: "", plate: "", phone: "", date: todayStr(), dept: "", purpose: "" });

  const handleSubmit = async (e) => {
    e.preventDefault();
    const extraLabel = reqType === "visitor" ? "소속팀" : "소속 부서";
    const extraValue = reqType === "visitor" ? form.purpose : form.dept;
    const checks = [
      ["이름", form.name],
      ["차량번호", form.plate],
      ["연락처", form.phone],
      [extraLabel, extraValue],
    ];
    const missing = checks.filter(([, value]) => !value || !value.trim()).map(([label]) => label);
    if (missing.length > 0) {
      setMissingFields(missing);
      return;
    }
    const newReq = {
      type: reqType,
      status: "pending",
      submittedAt: new Date().toISOString(),
      ...form,
    };
    try {
      await addDoc(collection(db, "requests"), newReq);
      resetForm();
      setConfirmation(newReq);
    } catch (error) {
      console.error("신청 저장 오류:", error);
      showToast("신청 접수 중 문제가 발생했어요. 다시 시도해주세요.");
    }
  };

  const goToAdmin = () => {
    if (adminUnlocked) {
      setView("admin");
      setUnseenCount(0);
    } else {
      setPasswordInput("");
      setPasswordError(false);
      setShowPasswordPrompt(true);
    }
  };

  const submitPassword = () => {
    if (passwordInput === ADMIN_PASSWORD) {
      setAdminUnlocked(true);
      setShowPasswordPrompt(false);
      setView("admin");
      setUnseenCount(0);
      setPasswordInput("");
      setPasswordError(false);
    } else {
      setPasswordError(true);
      setPasswordInput("");
    }
  };

  const updateStatus = async (id, status) => {
    try {
      await updateDoc(doc(db, "requests", id), { status });
    } catch (error) {
      console.error("상태 업데이트 오류:", error);
      showToast("처리 중 문제가 발생했어요. 다시 시도해주세요.");
    }
  };

  const deleteRequest = async (id) => {
    try {
      await deleteDoc(doc(db, "requests", id));
      setDeleteTarget(null);
      showToast("삭제됐어요.");
    } catch (error) {
      console.error("삭제 오류:", error);
      showToast("삭제 중 문제가 발생했어요. 다시 시도해주세요.");
    }
  };

  const filtered = useMemo(() => {
    if (filter === "all") return requests;
    return requests.filter((r) => r.status === filter);
  }, [requests, filter]);

  const counts = useMemo(() => {
    return requests.reduce(
      (acc, r) => {
        acc[r.status] = (acc[r.status] || 0) + 1;
        return acc;
      },
      { pending: 0, approved: 0, rejected: 0 }
    );
  }, [requests]);

  const lookupResults = useMemo(() => {
    const nameQ = lookupQuery.name.trim();
    const plateQ = lookupQuery.plate.trim();
    if (!nameQ && !plateQ) return [];
    return requests.filter((r) => {
      const nameMatch = nameQ ? r.name.includes(nameQ) : true;
      const plateMatch = plateQ ? r.plate.replace(/\s/g, "").includes(plateQ.replace(/\s/g, "")) : true;
      return nameMatch && plateMatch;
    });
  }, [requests, lookupQuery]);

  const exportExcel = () => {
    if (filtered.length === 0) {
      showToast("내보낼 신청 내역이 없어요.");
      return;
    }
    const rows = filtered.map((r) => ({
      구분: TYPE_META[r.type].label,
      이름: r.name,
      차량번호: r.plate,
      연락처: r.phone,
      날짜: r.date,
      "부서/목적": r.type === "visitor" ? r.purpose : r.dept,
      상태: STATUS_META[r.status].label,
      접수시각: new Date(r.submittedAt).toLocaleString("ko-KR"),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 14 },
      { wch: 12 }, { wch: 18 }, { wch: 8 }, { wch: 20 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "주차등록신청");
    XLSX.writeFile(wb, `주차등록신청_${new Date().toISOString().slice(0, 10)}.xlsx`);
    showToast("엑셀 파일을 내려받았어요.");
  };

  return (
    <div
      style={{
        fontFamily:
          "'Pretendard', 'Apple SD Gothic Neo', 'Noto Sans KR', -apple-system, sans-serif",
        background: "#F7F5EE",
        minHeight: "100%",
        color: "#1C1C1E",
        padding: "0",
      }}
    >
      {/* header */}
      <div
        style={{
          background: "#16233F",
          padding: "22px 24px 20px",
          borderBottom: "5px solid #F2A93B",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
          <div
            style={{
              background: "#FFFFFF",
              borderRadius: 8,
              padding: "7px 10px",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              flexShrink: 0,
            }}
          >
            <img src={LOGO_SRC} alt="IGM세계경영연구원" style={{ height: 20, display: "block" }} />
          </div>
          <h1
            style={{
              margin: 0,
              fontSize: 21,
              fontWeight: 800,
              color: "#FFFFFF",
              letterSpacing: "-0.01em",
            }}
          >
            주차 등록 요청
          </h1>
        </div>

        {/* notice */}
        <div
          style={{
            margin: "0 0 18px 0",
            background: "rgba(242,169,59,0.12)",
            border: "1px solid rgba(242,169,59,0.45)",
            borderRadius: 10,
            padding: "10px 14px",
            display: "flex",
            flexDirection: "column",
            gap: 4,
          }}
        >
          <span style={{ fontSize: 12, fontWeight: 800, color: "#F2A93B", letterSpacing: "0.02em" }}>
            ⚠ 주의사항
          </span>
          <span style={{ fontSize: 12.5, color: "#E8ECF5", lineHeight: 1.55 }}>
            1. 주차등록요청은 당일 오후 1시까지 접수받고 있으며, 이후 등록이 불가능합니다.<br />
            2. 오후 1시 이후 요청으로 미등록되어 주차요금이 부과될 경우, 개인비용으로 부과됩니다.<br />
            3. 주차등록은 오후 1시에 일괄등록됩니다. 이전에 출차하시는 분은 GA로 요청바랍니다.
          </span>
        </div>

        {/* tabs */}
        <div style={{ display: "flex", gap: 8 }}>
          <button
            onClick={() => setView("form")}
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "8px 14px", borderRadius: 8, border: "none",
              fontSize: 13, fontWeight: 700, cursor: "pointer",
              background: view === "form" ? "#F2A93B" : "rgba(255,255,255,0.08)",
              color: view === "form" ? "#16233F" : "#D8DEEA",
              transition: "all 0.15s",
            }}
          >
            <ClipboardList size={15} /> 신청하기
          </button>
          <button
            onClick={() => setView("lookup")}
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "8px 14px", borderRadius: 8, border: "none",
              fontSize: 13, fontWeight: 700, cursor: "pointer",
              background: view === "lookup" ? "#F2A93B" : "rgba(255,255,255,0.08)",
              color: view === "lookup" ? "#16233F" : "#D8DEEA",
              transition: "all 0.15s",
            }}
          >
            <Search size={15} /> 등록 확인
          </button>
          <button
            onClick={goToAdmin}
            style={{
              display: "flex", alignItems: "center", gap: 6,
              padding: "8px 14px", borderRadius: 8, border: "none",
              fontSize: 13, fontWeight: 700, cursor: "pointer",
              background: view === "admin" ? "#F2A93B" : "rgba(255,255,255,0.08)",
              color: view === "admin" ? "#16233F" : "#D8DEEA",
              transition: "all 0.15s",
              position: "relative",
            }}
          >
            <ShieldCheck size={15} /> 관리자 승인
            {!adminUnlocked && <Lock size={11} style={{ opacity: 0.8 }} />}
            {counts.pending > 0 && (
              <span
                style={{
                  background: "#C0392B", color: "#fff", fontSize: 10,
                  fontWeight: 800, borderRadius: 10, padding: "1px 6px",
                  marginLeft: 2,
                }}
              >
                {counts.pending}
              </span>
            )}
            {unseenCount > 0 && (
              <span
                style={{
                  position: "absolute", top: -5, right: -5,
                  width: 12, height: 12, borderRadius: "50%",
                  background: "#F2A93B", border: "2px solid #16233F",
                  animation: "igmPulse 1.4s infinite",
                }}
              />
            )}
          </button>
        </div>
      </div>

      {/* new-submission alert banner */}
      {unseenCount > 0 && view !== "admin" && (
        <div
          onClick={goToAdmin}
          style={{
            cursor: "pointer", background: "#FFF6E6", borderBottom: "1px solid #F2D89B",
            padding: "10px 24px", display: "flex", alignItems: "center", gap: 10,
            justifyContent: "center", fontSize: 13, fontWeight: 700, color: "#8A5A00",
          }}
        >
          <Bell size={15} style={{ animation: "igmPulse 1.4s infinite" }} />
          새로운 등록 신청이 {unseenCount}건 있어요 · 확인하러 가기 →
        </div>
      )}
      <style>{`
        @keyframes igmPulse {
          0% { box-shadow: 0 0 0 0 rgba(242,169,59,0.55); }
          70% { box-shadow: 0 0 0 6px rgba(242,169,59,0); }
          100% { box-shadow: 0 0 0 0 rgba(242,169,59,0); }
        }
      `}</style>

      <div style={{ padding: 24, maxWidth: 720, margin: "0 auto" }}>
        {loading ? (
          <div
            style={{
              textAlign: "center", padding: "60px 20px", color: "#A79F8A",
              fontSize: 13.5,
            }}
          >
            데이터를 불러오는 중이에요...
          </div>
        ) : (
        <>
        {view === "form" && (
          <div
            style={{
              background: "#FFFFFF", borderRadius: 14, padding: 22,
              border: "1px solid #ECE7D9", boxShadow: "0 1px 3px rgba(22,35,63,0.04)",
            }}
          >
            {/* type selector */}
            <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
              {Object.entries(TYPE_META).map(([key, meta]) => {
                const Icon = meta.icon;
                const active = reqType === key;
                return (
                  <button
                    key={key}
                    onClick={() => setReqType(key)}
                    style={{
                      flex: 1, display: "flex", alignItems: "center", justifyContent: "center",
                      gap: 8, padding: "12px", borderRadius: 10, cursor: "pointer",
                      border: active ? `2px solid ${meta.color}` : "2px solid #ECE7D9",
                      background: active ? `${meta.color}14` : "#FAF8F2",
                      fontWeight: 700, fontSize: 14,
                      color: active ? "#16233F" : "#8A8577",
                      transition: "all 0.15s",
                    }}
                  >
                    <Icon size={16} /> {meta.label} 등록
                  </button>
                );
              })}
            </div>

            <form onSubmit={handleSubmit} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
                <FormField label="이름 *">
                  <input style={inputStyle} value={form.name}
                    onChange={(e) => setForm({ ...form, name: e.target.value })}
                    placeholder="홍길동" />
                </FormField>
                <FormField label="차량번호 *">
                  <input style={inputStyle} value={form.plate}
                    onChange={(e) => setForm({ ...form, plate: e.target.value })}
                    placeholder="12가 3456" />
                </FormField>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 14 }}>
                <FormField label="연락처 *" hint="(숫자만 입력)">
                  <input
                    style={inputStyle}
                    value={form.phone}
                    onChange={(e) => setForm({ ...form, phone: formatPhone(e.target.value) })}
                    inputMode="numeric"
                    placeholder="010-0000-0000"
                  />
                </FormField>
                <FormField label={reqType === "visitor" ? "방문 날짜" : "등록 날짜"}>
                  <input
                    type="date"
                    style={{ ...inputStyle, background: "#F3F1E8", color: "#6B7280", cursor: "not-allowed" }}
                    value={form.date}
                    disabled
                    readOnly
                  />
                </FormField>
              </div>
              {reqType === "visitor" ? (
                <FormField label="소속팀 (커니) *">
                  <input style={inputStyle} value={form.purpose}
                    onChange={(e) => setForm({ ...form, purpose: e.target.value })}
                    placeholder="예: MS, SC, DT" />
                </FormField>
              ) : (
                <FormField label="소속 부서 (IGM) *">
                  <input style={inputStyle} value={form.dept}
                    onChange={(e) => setForm({ ...form, dept: e.target.value })}
                    placeholder="예: 마케팅팀" />
                </FormField>
              )}
              <button
                type="submit"
                style={{
                  marginTop: 6, padding: "12px", borderRadius: 9, border: "none",
                  background: "#16233F", color: "#fff", fontWeight: 700, fontSize: 14,
                  cursor: "pointer",
                }}
              >
                신청 접수하기
              </button>
            </form>
          </div>
        )}

        {view === "lookup" && (
          <div>
            <div
              style={{
                background: "#FFFFFF", borderRadius: 14, padding: 20,
                border: "1px solid #ECE7D9", marginBottom: 14,
                display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 12,
              }}
            >
              <FormField label="이름">
                <input
                  style={inputStyle}
                  value={lookupQuery.name}
                  onChange={(e) => setLookupQuery({ ...lookupQuery, name: e.target.value })}
                  placeholder="홍길동"
                />
              </FormField>
              <FormField label="차량번호">
                <input
                  style={inputStyle}
                  value={lookupQuery.plate}
                  onChange={(e) => setLookupQuery({ ...lookupQuery, plate: e.target.value })}
                  placeholder="12가 3456"
                />
              </FormField>
            </div>

            {!lookupQuery.name.trim() && !lookupQuery.plate.trim() ? (
              <div
                style={{
                  textAlign: "center", padding: "50px 20px", color: "#A79F8A",
                  background: "#FFFFFF", borderRadius: 14, border: "1px dashed #E4DFD1",
                  fontSize: 13.5,
                }}
              >
                이름 또는 차량번호를 입력하면 신청 처리 상태를 확인할 수 있어요.
              </div>
            ) : lookupResults.length === 0 ? (
              <div
                style={{
                  textAlign: "center", padding: "50px 20px", color: "#A79F8A",
                  background: "#FFFFFF", borderRadius: 14, border: "1px dashed #E4DFD1",
                  fontSize: 13.5,
                }}
              >
                일치하는 신청 내역이 없어요.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {lookupResults.map((r) => {
                  const type = TYPE_META[r.type];
                  const status = STATUS_META[r.status];
                  const Icon = type.icon;
                  return (
                    <div
                      key={r.id}
                      style={{
                        background: "#fff", borderRadius: 12, padding: 14,
                        border: "1px solid #ECE7D9", display: "flex",
                        alignItems: "center", gap: 14,
                      }}
                    >
                      <GateIcon status={r.status} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
                          <Icon size={13} color={type.color} />
                          <span style={{ fontSize: 12, fontWeight: 700, color: type.color }}>
                            {type.label}
                          </span>
                        </div>
                        <div style={{ fontSize: 14.5, fontWeight: 700 }}>
                          {r.name} <span style={{ color: "#8A8577", fontWeight: 500 }}>· {r.plate}</span>
                        </div>
                        <div style={{ fontSize: 12.5, color: "#8A8577", marginTop: 2 }}>
                          {r.date} {r.phone && `· ${r.phone}`}
                        </div>
                      </div>
                      <span
                        style={{
                          fontSize: 12, fontWeight: 800, padding: "5px 12px",
                          borderRadius: 20, background: status.bg, color: status.color,
                          flexShrink: 0,
                        }}
                      >
                        {status.label}
                      </span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {view === "admin" && (
          <div>
            <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
              {["all", "pending", "approved", "rejected"].map((f) => (
                <button
                  key={f}
                  onClick={() => setFilter(f)}
                  style={{
                    padding: "7px 13px", borderRadius: 20, fontSize: 12.5, fontWeight: 700,
                    cursor: "pointer", border: "1.5px solid #E4DFD1",
                    background: filter === f ? "#16233F" : "#fff",
                    color: filter === f ? "#fff" : "#5B5646",
                  }}
                >
                  {f === "all" ? "전체" : STATUS_META[f].label}
                  {f !== "all" && ` (${counts[f] || 0})`}
                </button>
              ))}
              <button
                onClick={exportExcel}
                style={{
                  marginLeft: "auto", display: "flex", alignItems: "center", gap: 6,
                  padding: "7px 13px", borderRadius: 20, fontSize: 12.5, fontWeight: 700,
                  cursor: "pointer", border: "1.5px solid #2F7D52", background: "#E4F3EA",
                  color: "#2F7D52",
                }}
              >
                <Download size={13} /> 엑셀로 다운로드
              </button>
            </div>

            {filtered.length === 0 ? (
              <div
                style={{
                  textAlign: "center", padding: "50px 20px", color: "#A79F8A",
                  background: "#FFFFFF", borderRadius: 14, border: "1px dashed #E4DFD1",
                  fontSize: 13.5,
                }}
              >
                아직 신청 내역이 없어요. '신청하기' 탭에서 등록해보세요.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {filtered.map((r) => {
                  const type = TYPE_META[r.type];
                  const status = STATUS_META[r.status];
                  const Icon = type.icon;
                  return (
                    <div
                      key={r.id}
                      style={{
                        background: "#fff", borderRadius: 12, padding: 14,
                        border: "1px solid #ECE7D9", display: "flex",
                        alignItems: "center", gap: 14,
                      }}
                    >
                      <GateIcon status={r.status} />
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
                          <Icon size={13} color={type.color} />
                          <span style={{ fontSize: 12, fontWeight: 700, color: type.color }}>
                            {type.label}
                          </span>
                          <span
                            style={{
                              fontSize: 11, fontWeight: 700, padding: "1px 8px",
                              borderRadius: 10, background: status.bg, color: status.color,
                            }}
                          >
                            {status.label}
                          </span>
                        </div>
                        <div style={{ fontSize: 14.5, fontWeight: 700 }}>
                          {r.name} <span style={{ color: "#8A8577", fontWeight: 500 }}>· {r.plate}</span>
                        </div>
                        <div style={{ fontSize: 12.5, color: "#8A8577", marginTop: 2 }}>
                          {r.date} {r.phone && `· ${r.phone}`}
                          {r.type === "visitor" && r.purpose ? ` · ${r.purpose}` : ""}
                          {r.type === "employee" && r.dept ? ` · ${r.dept}` : ""}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                        {r.status === "pending" ? (
                          <>
                            <button
                              onClick={() => updateStatus(r.id, "approved")}
                              title="승인"
                              style={{
                                width: 32, height: 32, borderRadius: 8, border: "none",
                                background: "#2F7D52", color: "#fff", cursor: "pointer",
                                display: "flex", alignItems: "center", justifyContent: "center",
                              }}
                            >
                              <Check size={16} />
                            </button>
                            <button
                              onClick={() => updateStatus(r.id, "rejected")}
                              title="거절"
                              style={{
                                width: 32, height: 32, borderRadius: 8, border: "none",
                                background: "#C0392B", color: "#fff", cursor: "pointer",
                                display: "flex", alignItems: "center", justifyContent: "center",
                              }}
                            >
                              <X size={16} />
                            </button>
                          </>
                        ) : (
                          <button
                            onClick={() => updateStatus(r.id, "pending")}
                            title="대기로 되돌리기"
                            style={{
                              width: 32, height: 32, borderRadius: 8,
                              border: "1.5px solid #E4DFD1", background: "#fff",
                              color: "#8A8577", cursor: "pointer",
                              display: "flex", alignItems: "center", justifyContent: "center",
                            }}
                          >
                            <Clock size={15} />
                          </button>
                        )}
                        <button
                          onClick={() => setDeleteTarget(r)}
                          title="삭제"
                          style={{
                            width: 32, height: 32, borderRadius: 8,
                            border: "1.5px solid #F5D9D2", background: "#FFF6F4",
                            color: "#C0392B", cursor: "pointer",
                            display: "flex", alignItems: "center", justifyContent: "center",
                          }}
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
        </>
        )}
      </div>

      {toast && (
        <div
          style={{
            position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)",
            background: "#16233F", color: "#fff", padding: "10px 18px",
            borderRadius: 10, fontSize: 13, fontWeight: 600, boxShadow: "0 4px 14px rgba(0,0,0,0.2)",
            zIndex: 50,
          }}
        >
          {toast}
        </div>
      )}

      {missingFields && (
        <div
          style={{
            position: "fixed", inset: 0, background: "rgba(22,35,63,0.55)",
            display: "flex", alignItems: "center", justifyContent: "center",
            zIndex: 100, padding: 20,
          }}
          onClick={() => setMissingFields(null)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "#fff", borderRadius: 16, padding: "28px 26px",
              maxWidth: 340, width: "100%", textAlign: "center",
              boxShadow: "0 12px 32px rgba(22,35,63,0.25)",
            }}
          >
            <div
              style={{
                width: 56, height: 56, borderRadius: "50%", background: "#FBEAE7",
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 14px",
              }}
            >
              <X size={28} color="#C0392B" strokeWidth={3} />
            </div>
            <h2 style={{ margin: "0 0 8px", fontSize: 17, fontWeight: 800, color: "#16233F" }}>
              입력하지 않은 항목이 있어요
            </h2>
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: "#8A8577", lineHeight: 1.6 }}>
              아래 항목을 모두 입력해야 신청할 수 있어요.
            </p>
            <div
              style={{
                background: "#FAF8F2", borderRadius: 10, padding: "12px 16px",
                textAlign: "left", marginBottom: 20,
              }}
            >
              {missingFields.map((label) => (
                <div
                  key={label}
                  style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#C0392B", padding: "4px 0" }}
                >
                  <span style={{ width: 5, height: 5, borderRadius: "50%", background: "#C0392B", flexShrink: 0 }} />
                  {label}
                </div>
              ))}
            </div>
            <button
              onClick={() => setMissingFields(null)}
              style={{
                width: "100%", padding: "11px", borderRadius: 9, border: "none",
                background: "#16233F", color: "#fff", fontWeight: 700, fontSize: 14,
                cursor: "pointer",
              }}
            >
              확인
            </button>
          </div>
        </div>
      )}

      {showPasswordPrompt && (
        <div
          style={{
            position: "fixed", inset: 0, background: "rgba(22,35,63,0.55)",
            display: "flex", alignItems: "center", justifyContent: "center",
            zIndex: 100, padding: 20,
          }}
          onClick={() => setShowPasswordPrompt(false)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "#fff", borderRadius: 16, padding: "28px 26px",
              maxWidth: 320, width: "100%", textAlign: "center",
              boxShadow: "0 12px 32px rgba(22,35,63,0.25)",
            }}
          >
            <div
              style={{
                width: 56, height: 56, borderRadius: "50%", background: "#EAEEF6",
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 14px",
              }}
            >
              <Lock size={26} color="#16233F" strokeWidth={2.5} />
            </div>
            <h2 style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 800, color: "#16233F" }}>
              관리자 비밀번호
            </h2>
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: "#8A8577" }}>
              4자리 비밀번호를 입력해주세요.
            </p>
            <input
              type="password"
              inputMode="numeric"
              maxLength={4}
              autoFocus
              value={passwordInput}
              onChange={(e) => {
                setPasswordInput(e.target.value.replace(/[^0-9]/g, "").slice(0, 4));
                setPasswordError(false);
              }}
              onKeyDown={(e) => e.key === "Enter" && submitPassword()}
              style={{
                width: "100%", boxSizing: "border-box", textAlign: "center",
                fontSize: 22, letterSpacing: "0.5em", padding: "12px 10px",
                borderRadius: 9, border: passwordError ? "1.5px solid #C0392B" : "1.5px solid #E4DFD1",
                marginBottom: passwordError ? 8 : 18, outline: "none",
              }}
              placeholder="••••"
            />
            {passwordError && (
              <p style={{ margin: "0 0 14px", fontSize: 12, color: "#C0392B", fontWeight: 700 }}>
                비밀번호가 올바르지 않아요.
              </p>
            )}
            <button
              onClick={submitPassword}
              style={{
                width: "100%", padding: "11px", borderRadius: 9, border: "none",
                background: "#16233F", color: "#fff", fontWeight: 700, fontSize: 14,
                cursor: "pointer",
              }}
            >
              확인
            </button>
          </div>
        </div>
      )}

      {deleteTarget && (
        <div
          style={{
            position: "fixed", inset: 0, background: "rgba(22,35,63,0.55)",
            display: "flex", alignItems: "center", justifyContent: "center",
            zIndex: 100, padding: 20,
          }}
          onClick={() => setDeleteTarget(null)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "#fff", borderRadius: 16, padding: "28px 26px",
              maxWidth: 340, width: "100%", textAlign: "center",
              boxShadow: "0 12px 32px rgba(22,35,63,0.25)",
            }}
          >
            <div
              style={{
                width: 56, height: 56, borderRadius: "50%", background: "#FBEAE7",
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 14px",
              }}
            >
              <Trash2 size={26} color="#C0392B" strokeWidth={2.5} />
            </div>
            <h2 style={{ margin: "0 0 8px", fontSize: 17, fontWeight: 800, color: "#16233F" }}>
              이 신청을 삭제할까요?
            </h2>
            <p style={{ margin: "0 0 16px", fontSize: 12.5, color: "#8A8577", lineHeight: 1.6 }}>
              삭제하면 되돌릴 수 없어요. 필요하면 삭제 전에 엑셀로 먼저 백업해두세요.
            </p>
            <div
              style={{
                background: "#FAF8F2", borderRadius: 10, padding: "12px 16px",
                textAlign: "left", marginBottom: 20, fontSize: 13,
              }}
            >
              <div style={{ fontWeight: 700 }}>
                {deleteTarget.name} <span style={{ color: "#8A8577", fontWeight: 500 }}>· {deleteTarget.plate}</span>
              </div>
              <div style={{ color: "#8A8577", marginTop: 2 }}>
                {TYPE_META[deleteTarget.type].label} · {deleteTarget.date}
              </div>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button
                onClick={() => setDeleteTarget(null)}
                style={{
                  flex: 1, padding: "11px", borderRadius: 9, border: "1.5px solid #E4DFD1",
                  background: "#fff", color: "#5B5646", fontWeight: 700, fontSize: 14,
                  cursor: "pointer",
                }}
              >
                취소
              </button>
              <button
                onClick={() => deleteRequest(deleteTarget.id)}
                style={{
                  flex: 1, padding: "11px", borderRadius: 9, border: "none",
                  background: "#C0392B", color: "#fff", fontWeight: 700, fontSize: 14,
                  cursor: "pointer",
                }}
              >
                삭제
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmation && (
        <div
          style={{
            position: "fixed", inset: 0, background: "rgba(22,35,63,0.55)",
            display: "flex", alignItems: "center", justifyContent: "center",
            zIndex: 100, padding: 20,
          }}
          onClick={() => setConfirmation(null)}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              background: "#fff", borderRadius: 16, padding: "28px 26px",
              maxWidth: 360, width: "100%", textAlign: "center",
              boxShadow: "0 12px 32px rgba(22,35,63,0.25)",
            }}
          >
            <div
              style={{
                width: 56, height: 56, borderRadius: "50%", background: "#E4F3EA",
                display: "flex", alignItems: "center", justifyContent: "center",
                margin: "0 auto 14px",
              }}
            >
              <Check size={28} color="#2F7D52" strokeWidth={3} />
            </div>
            <h2 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 800, color: "#16233F" }}>
              등록이 완료되었습니다
            </h2>
            <p style={{ margin: "0 0 18px", fontSize: 12.5, color: "#8A8577" }}>
              아래 내용으로 주차등록 신청이 접수됐어요.
            </p>
            <div
              style={{
                background: "#FAF8F2", borderRadius: 10, padding: "14px 16px",
                textAlign: "left", display: "flex", flexDirection: "column", gap: 8,
                marginBottom: 20,
              }}
            >
              {[
                ["구분", TYPE_META[confirmation.type].label],
                ["이름", confirmation.name],
                ["차량번호", confirmation.plate],
                ["연락처", confirmation.phone || "-"],
                ["날짜", confirmation.date],
                [
                  confirmation.type === "visitor" ? "소속팀" : "소속 부서",
                  (confirmation.type === "visitor" ? confirmation.purpose : confirmation.dept) || "-",
                ],
              ].map(([label, value]) => (
                <div key={label} style={{ display: "flex", justifyContent: "space-between", fontSize: 13 }}>
                  <span style={{ color: "#8A8577" }}>{label}</span>
                  <span style={{ fontWeight: 700, color: "#1C1C1E" }}>{value}</span>
                </div>
              ))}
            </div>
            <button
              onClick={() => setConfirmation(null)}
              style={{
                width: "100%", padding: "11px", borderRadius: 9, border: "none",
                background: "#16233F", color: "#fff", fontWeight: 700, fontSize: 14,
                cursor: "pointer",
              }}
            >
              확인
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
