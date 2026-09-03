import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';

// Utility: convert column index to Excel letters (0 -> A, 1 -> B, 26 -> AA)
const getColLetter = (index) => {
  let letter = '';
  let temp = index;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
};

export default function App() {
  // ---------------- AUTHENTICATION STATE ----------------
  const [currentUser, setCurrentUser] = useState(null); // { email, role: 'admin' | 'user', name: string }
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authError, setAuthError] = useState('');

  // ---------------- WORKSPACE DATA STATE ----------------
  const [workbookData, setWorkbookData] = useState(null);
  const [activeSheetName, setActiveSheetName] = useState('');
  const [activeCell, setActiveCell] = useState([0, 0]);
  const [selectedRowIndices, setSelectedRowIndices] = useState([]);

  // Top Ribbon & Filtering Controls
  const [quickView, setQuickView] = useState('All Records');
  const [syncTime, setSyncTime] = useState('12:00 PM');
  const [searchQuery, setSearchQuery] = useState('');
  const [filters, setFilters] = useState({ date: '', stage: '', owner: '', campaign: '' });

  // Modals & Preferences
  const [moreFiltersModalOpen, setMoreFiltersModalOpen] = useState(false);
  const [selectedCustomCol, setSelectedCustomCol] = useState('');
  const [customColValue, setCustomColValue] = useState('');
  const [settingsModalOpen, setSettingsModalOpen] = useState(false);
  const [autoTrimOnImport, setAutoTrimOnImport] = useState(true);
  const [confirmBeforeDelete, setConfirmBeforeDelete] = useState(true);

  const fileInputRef = useRef(null);

  // ---------------- AUTH HANDLERS ----------------
  const handleFormLogin = (e) => {
    e.preventDefault();
    setAuthError('');

    if (!authEmail.trim() || !authPassword.trim()) {
      setAuthError('Please fill in both email and password.');
      return;
    }

    if (authEmail.toLowerCase().includes('admin')) {
      setCurrentUser({
        email: authEmail,
        role: 'admin',
        name: 'Master Administrator'
      });
    } else {
      setCurrentUser({
        email: authEmail,
        role: 'user',
        name: 'Counselor Priya Sharma'
      });
    }
  };

  const handleQuickLogin = (role) => {
    setAuthError('');
    if (role === 'admin') {
      setCurrentUser({
        email: 'admin@lms.edu',
        role: 'admin',
        name: 'Master Administrator'
      });
    } else {
      setCurrentUser({
        email: 'priya.sharma@lms.edu',
        role: 'user',
        name: 'Counselor Priya Sharma'
      });
    }
  };

  const handleSignOut = () => {
    setCurrentUser(null);
    setSelectedRowIndices([]);
    setSearchQuery('');
    setAuthEmail('');
    setAuthPassword('');
  };

  // ---------------- SPREADSHEET PARSER ----------------
  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target.result;
        const wb = XLSX.read(buffer, { type: 'array' });
        const sheetNames = wb.SheetNames;
        const parsedSheets = {};

        sheetNames.forEach((name) => {
          const matrix = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
          const firstRow = matrix.findIndex((r) => r.filter((c) => String(c).trim() !== '').length >= 2);
          let cleanMatrix = firstRow !== -1 ? matrix.slice(firstRow) : matrix;

          if (autoTrimOnImport) {
            cleanMatrix = cleanMatrix.map((row) =>
              row.map((cell) => (typeof cell === 'string' ? cell.trim() : cell))
            );
          }

          parsedSheets[name] = cleanMatrix.length > 0 ? cleanMatrix : [['']];
        });

        setWorkbookData({
          fileName: file.name,
          sheetNames,
          sheets: parsedSheets,
        });
        setActiveSheetName(sheetNames[0]);
        setActiveCell([0, 0]);
        setSelectedRowIndices([]);
        setSearchQuery('');
        setFilters({ date: '', stage: '', owner: '', campaign: '' });
        setSyncTime(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }));
      } catch (err) {
        console.error('File parsing error:', err);
        alert('Could not read the spreadsheet. Please verify it is a valid .xlsx or .csv file.');
      } finally {
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExport = () => {
    if (!workbookData || !activeSheetName) return;
    const currentRows = workbookData.sheets[activeSheetName] || [];
    const worksheet = XLSX.utils.aoa_to_sheet(currentRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, worksheet, activeSheetName);
    XLSX.writeFile(wb, `LMS_Export_${workbookData.fileName || 'Leads.xlsx'}`);
  };

  const handleResetFilters = () => {
    setSearchQuery('');
    setQuickView('All Records');
    setFilters({ date: '', stage: '', owner: '', campaign: '' });
    setSelectedCustomCol('');
    setCustomColValue('');
  };

  // ---------------- ROLE & FILTER PIPELINE ----------------
  const currentSheetRows = useMemo(() => {
    if (!workbookData || !activeSheetName) return [];
    const raw = workbookData.sheets[activeSheetName] || [];
    if (raw.length <= 1) return raw;

    const headers = raw[0];
    let dataRows = raw.slice(1);

    const findColIdx = (regex) => headers.findIndex((h) => regex.test(String(h)));
    const dateIdx = findColIdx(/date/i);
    const stageIdx = findColIdx(/stage|status/i);
    const ownerIdx = findColIdx(/owner|counselor|team/i);
    const campaignIdx = findColIdx(/campaign|source/i);

    // Row-level Security: Counselors only see their own assigned leads
    if (currentUser?.role === 'user' && ownerIdx !== -1) {
      dataRows = dataRows.filter((r) =>
        String(r[ownerIdx] ?? '').toLowerCase().includes(currentUser.name.toLowerCase()) ||
        String(r[ownerIdx] ?? '').toLowerCase().includes('priya')
      );
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      dataRows = dataRows.filter((r) => r.some((c) => String(c ?? '').toLowerCase().includes(q)));
    }

    if (filters.date && dateIdx !== -1) {
      dataRows = dataRows.filter((r) => String(r[dateIdx] ?? '').includes(filters.date));
    }
    if (filters.stage && stageIdx !== -1) {
      dataRows = dataRows.filter((r) => String(r[stageIdx] ?? '').toLowerCase() === filters.stage.toLowerCase());
    }
    if (filters.owner && ownerIdx !== -1) {
      dataRows = dataRows.filter((r) => String(r[ownerIdx] ?? '').toLowerCase() === filters.owner.toLowerCase());
    }
    if (filters.campaign && campaignIdx !== -1) {
      dataRows = dataRows.filter((r) => String(r[campaignIdx] ?? '').toLowerCase() === filters.campaign.toLowerCase());
    }

    if (selectedCustomCol !== '' && customColValue.trim()) {
      const cIdx = Number(selectedCustomCol);
      const valQuery = customColValue.toLowerCase();
      dataRows = dataRows.filter((r) => String(r[cIdx] ?? '').toLowerCase().includes(valQuery));
    }

    if (quickView === 'Qualified') {
      dataRows = dataRows.filter((r) => r.some((c) => /qualified|fee paid/i.test(String(c ?? ''))));
    } else if (quickView === 'New Leads') {
      dataRows = dataRows.filter((r) => r.some((c) => /new lead|inquiry/i.test(String(c ?? ''))));
    }

    return [headers, ...dataRows];
  }, [workbookData, activeSheetName, searchQuery, filters, selectedCustomCol, customColValue, quickView, currentUser]);

  const totalCols = useMemo(() => {
    if (currentSheetRows.length === 0) return 0;
    return Math.max(...currentSheetRows.map((r) => r.length));
  }, [currentSheetRows]);

  // Checkbox Batch Selection Handlers
  const handleToggleSelectAll = () => {
    if (selectedRowIndices.length === currentSheetRows.length - 1 && currentSheetRows.length > 1) {
      setSelectedRowIndices([]);
    } else {
      const allDataIndices = currentSheetRows.slice(1).map((_, idx) => idx + 1);
      setSelectedRowIndices(allDataIndices);
    }
  };

  const handleToggleRow = (rIdx) => {
    if (rIdx === 0) return;
    setSelectedRowIndices((prev) =>
      prev.includes(rIdx) ? prev.filter((i) => i !== rIdx) : [...prev, rIdx]
    );
  };

  const handleDeleteCheckedRows = () => {
    if (!workbookData || !activeSheetName || selectedRowIndices.length === 0) return;
    if (currentUser?.role !== 'admin') {
      alert('Unauthorized: Only administrators can delete lead rows.');
      return;
    }

    if (confirmBeforeDelete && !window.confirm(`Delete ${selectedRowIndices.length} selected row(s)?`)) {
      return;
    }

    const currentRows = workbookData.sheets[activeSheetName] || [];
    const updatedRows = currentRows.filter((_, idx) => !selectedRowIndices.includes(idx));

    setWorkbookData((prev) => ({
      ...prev,
      sheets: { ...prev.sheets, [activeSheetName]: updatedRows },
    }));
    setSelectedRowIndices([]);
    setActiveCell([0, 0]);
  };

  // DD-MM-YYYY Date Parser
  const getUniqueMonthYearOptions = () => {
    if (!workbookData || !activeSheetName) return [];
    const matrix = workbookData.sheets[activeSheetName] || [];
    if (matrix.length <= 1) return [];

    const dateIdx = matrix[0].findIndex((h) => /date/i.test(String(h)));
    if (dateIdx === -1) return [];

    const monthNames = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];

    const optionsMap = new Map();
    matrix.slice(1).forEach((row) => {
      const rawVal = String(row[dateIdx] ?? '').trim();
      const parts = rawVal.split(/[-/]/);
      if (parts.length === 3 && parts[1] && parts[2]) {
        const monthNum = parseInt(parts[1], 10);
        const year = parts[2];
        if (monthNum >= 1 && monthNum <= 12) {
          const label = `${monthNames[monthNum - 1]} ${year}`;
          const filterKey = `-${String(monthNum).padStart(2, '0')}-${year}`;
          optionsMap.set(filterKey, label);
        }
      }
    });

    return Array.from(optionsMap.entries()).map(([value, label]) => ({ value, label }));
  };

  const getUniqueOptions = (regex) => {
    if (!workbookData || !activeSheetName) return [];
    const matrix = workbookData.sheets[activeSheetName] || [];
    if (matrix.length <= 1) return [];
    const idx = matrix[0].findIndex((h) => regex.test(String(h)));
    if (idx === -1) return [];
    return Array.from(new Set(matrix.slice(1).map((r) => String(r[idx] ?? '').trim()))).filter(Boolean);
  };

  const activeCellValue = currentSheetRows[activeCell[0]]?.[activeCell[1]] ?? '';
  const isAdmin = currentUser?.role === 'admin';

  // =========================================================================
  // VIEW 1: AUTHENTICATION GATEWAY
  // =========================================================================
  if (!currentUser) {
    return (
      <div style={{ height: '100vh', width: '100vw', display: 'flex', alignItems: 'center', justifyContent: 'center', backgroundColor: '#0f172a', fontFamily: 'Segoe UI, sans-serif' }}>
        <div style={{ backgroundColor: '#ffffff', padding: '36px', borderRadius: '12px', width: '380px', boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.4)' }}>
          <div style={{ textAlign: 'center', marginBottom: '24px' }}>
            <div style={{ width: '48px', height: '48px', backgroundColor: '#2563eb', borderRadius: '10px', color: '#fff', fontSize: '1.1rem', fontWeight: 'bold', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 12px' }}>
              LMS
            </div>
            <h2 style={{ margin: '0 0 4px', fontSize: '1.3rem', color: '#0f172a' }}>Leads Management System</h2>
            <span style={{ fontSize: '0.82rem', color: '#64748b' }}>Select a persona or enter your work email</span>
          </div>

          {authError && (
            <div style={{ padding: '8px 12px', backgroundColor: '#fef2f2', border: '1px solid #fecaca', borderRadius: '6px', color: '#dc2626', fontSize: '0.8rem', marginBottom: '14px' }}>
              {authError}
            </div>
          )}

          <form onSubmit={handleFormLogin} style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 'bold', color: '#475569', marginBottom: '4px' }}>Work Email</label>
              <input
                type="email"
                placeholder="admin@lms.edu or counselor@lms.edu"
                value={authEmail}
                onChange={(e) => setAuthEmail(e.target.value)}
                style={{ width: '100%', padding: '8px 12px', border: '1px solid #cbd5e1', borderRadius: '6px', fontSize: '0.85rem', boxSizing: 'border-box' }}
              />
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 'bold', color: '#475569', marginBottom: '4px' }}>Password</label>
              <input
                type="password"
                placeholder="••••••••••••"
                value={authPassword}
                onChange={(e) => setAuthPassword(e.target.value)}
                style={{ width: '100%', padding: '8px 12px', border: '1px solid #cbd5e1', borderRadius: '6px', fontSize: '0.85rem', boxSizing: 'border-box' }}
              />
            </div>

            <button
              type="submit"
              style={{ backgroundColor: '#2563eb', color: '#fff', border: 'none', padding: '10px', borderRadius: '6px', fontWeight: 'bold', fontSize: '0.88rem', cursor: 'pointer', marginTop: '4px' }}
            >
              Sign In
            </button>
          </form>

          <div style={{ marginTop: '22px', paddingTop: '18px', borderTop: '1px solid #e2e8f0', textAlign: 'center' }}>
            <span style={{ fontSize: '0.72rem', color: '#94a3b8', textTransform: 'uppercase', fontWeight: 'bold', letterSpacing: '0.5px' }}>
              Instant Demo Roles
            </span>
            <div style={{ display: 'flex', gap: '8px', marginTop: '10px' }}>
              <button
                type="button"
                onClick={() => handleQuickLogin('admin')}
                style={{ flex: 1, padding: '8px 10px', backgroundColor: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px', color: '#1d4ed8', fontSize: '0.76rem', fontWeight: 'bold', cursor: 'pointer' }}
              >
                Login as Admin
              </button>
              <button
                type="button"
                onClick={() => handleQuickLogin('user')}
                style={{ flex: 1, padding: '8px 10px', backgroundColor: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '6px', color: '#334155', fontSize: '0.76rem', fontWeight: 'bold', cursor: 'pointer' }}
              >
                Login as Counselor
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // =========================================================================
  // VIEW 2: LEADS DESK WORKSPACE
  // =========================================================================
  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw', backgroundColor: '#f8fafc', fontFamily: 'Segoe UI, -apple-system, sans-serif', overflow: 'hidden' }}>
      
      {/* SIDEBAR */}
      <aside style={{ width: '230px', backgroundColor: '#0f172a', color: '#f8fafc', display: 'flex', flexDirection: 'column', borderRight: '1px solid #1e293b' }}>
        <div style={{ padding: '18px 16px', borderBottom: '1px solid #1e293b', display: 'flex', alignItems: 'center', gap: '10px' }}>
          <div style={{ width: '32px', height: '32px', borderRadius: '8px', backgroundColor: '#2563eb', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 'bold', fontSize: '0.85rem', color: '#ffffff' }}>
            LMS
          </div>
          <div>
            <span style={{ fontSize: '0.92rem', fontWeight: 'bold', color: '#ffffff', display: 'block' }}>Leads Management</span>
            <span style={{ fontSize: '0.72rem', color: '#94a3b8' }}>System Workspace</span>
          </div>
        </div>

        <nav style={{ flex: 1, padding: '12px 8px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
          <button
            style={{
              display: 'block',
              padding: '10px 14px',
              borderRadius: '6px',
              border: 'none',
              backgroundColor: '#2563eb',
              color: '#ffffff',
              fontWeight: '600',
              cursor: 'pointer',
              textAlign: 'left',
              fontSize: '0.84rem',
            }}
          >
            Leads Master Desk
          </button>

          {isAdmin && (
            <button
              onClick={() => setSettingsModalOpen(true)}
              style={{
                display: 'block',
                padding: '10px 14px',
                borderRadius: '6px',
                border: 'none',
                backgroundColor: 'transparent',
                color: '#94a3b8',
                fontWeight: '500',
                cursor: 'pointer',
                textAlign: 'left',
                fontSize: '0.84rem',
              }}
            >
              Settings
            </button>
          )}
        </nav>

        {/* LOGGED IN USER CARD */}
        <div style={{ padding: '14px 16px', borderTop: '1px solid #1e293b', backgroundColor: '#0b1329', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <span style={{ display: 'block', fontSize: '0.8rem', fontWeight: 'bold', color: '#fff' }}>{currentUser.name}</span>
            <span style={{ display: 'block', fontSize: '0.68rem', color: isAdmin ? '#38bdf8' : '#a7f3d0' }}>
              {isAdmin ? '● Full Administrator' : '● Counselor Desk'}
            </span>
          </div>
          <button
            onClick={handleSignOut}
            title="Sign Out"
            style={{ background: 'none', border: '1px solid #334155', color: '#94a3b8', padding: '3px 8px', borderRadius: '4px', cursor: 'pointer', fontSize: '0.72rem' }}
          >
            Sign Out
          </button>
        </div>
      </aside>

      {/* MAIN WORKSPACE */}
      <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
        <input type="file" ref={fileInputRef} accept=".xlsx, .xls, .csv" onChange={handleFileChange} style={{ display: 'none' }} />

        {/* TOP TOOLBAR */}
        <header style={{ backgroundColor: '#ffffff', borderBottom: '1px solid #e2e8f0', padding: '10px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '10px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
            <span style={{ fontSize: '1.1rem', fontWeight: '800', color: '#0f172a' }}>
              {isAdmin ? 'Leads Management System (Admin)' : 'Counselor Lead Desk'}
            </span>

            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', borderLeft: '1px solid #e2e8f0', paddingLeft: '12px' }}>
              <span style={{ fontSize: '0.78rem', color: '#64748b' }}>Quick View:</span>
              <select
                value={quickView}
                onChange={(e) => setQuickView(e.target.value)}
                style={{
                  backgroundColor: '#eff6ff',
                  border: '1px solid #bfdbfe',
                  color: '#2563eb',
                  fontWeight: '600',
                  fontSize: '0.8rem',
                  padding: '3px 8px',
                  borderRadius: '5px',
                  outline: 'none',
                }}
              >
                <option value="All Records">All Records</option>
                <option value="Qualified">Qualified Prospects</option>
                <option value="New Leads">New Inquiries</option>
              </select>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', backgroundColor: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '5px', padding: '3px 8px', fontSize: '0.75rem', color: '#64748b' }}>
              <span>Last sync on:</span>
              <strong style={{ color: '#0f172a' }}>{syncTime}</strong>
              <span
                onClick={() => setSyncTime(new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }))}
                title="Sync Now"
                style={{ cursor: 'pointer', color: '#2563eb', marginLeft: '4px' }}
              >
                ↻
              </span>
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <div style={{ position: 'relative' }}>
              <input
                type="text"
                placeholder="Search lead records..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                disabled={!workbookData}
                style={{
                  width: '180px',
                  padding: '5px 8px 5px 12px',
                  borderRadius: '5px',
                  border: '1px solid #cbd5e1',
                  fontSize: '0.8rem',
                  outline: 'none',
                }}
              />
            </div>

            {isAdmin && (
              <button
                onClick={() => fileInputRef.current?.click()}
                style={{ backgroundColor: '#0f172a', color: '#fff', border: 'none', padding: '6px 12px', borderRadius: '5px', fontSize: '0.8rem', fontWeight: '600', cursor: 'pointer' }}
              >
                Import Excel File
              </button>
            )}

            {isAdmin && (
              <button
                onClick={handleExport}
                disabled={!workbookData}
                style={{ backgroundColor: workbookData ? '#ffffff' : '#f8fafc', color: workbookData ? '#0f172a' : '#94a3b8', border: '1px solid #cbd5e1', padding: '6px 12px', borderRadius: '5px', fontSize: '0.8rem', cursor: workbookData ? 'pointer' : 'not-allowed' }}
              >
                Export Leads (.xlsx)
              </button>
            )}
          </div>
        </header>

        {/* FILTER RIBBON */}
        <div style={{ backgroundColor: '#ffffff', borderBottom: '1px solid #e2e8f0', padding: '8px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '10px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
            <span style={{ fontSize: '0.78rem', color: '#64748b', fontWeight: 'bold' }}>Filters:</span>

            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <span style={{ fontSize: '0.65rem', color: '#64748b', fontWeight: 'bold' }}>Registration Date</span>
              <select
                value={filters.date}
                onChange={(e) => setFilters({ ...filters, date: e.target.value })}
                style={{ border: '1px solid #cbd5e1', borderRadius: '4px', padding: '3px 6px', fontSize: '0.76rem', backgroundColor: '#fff', color: '#334155' }}
              >
                <option value="">Select Here</option>
                {getUniqueMonthYearOptions().length > 0 ? (
                  getUniqueMonthYearOptions().map((opt) => (
                    <option key={opt.value} value={opt.value}>
                      {opt.label}
                    </option>
                  ))
                ) : (
                  <>
                    <option value="-07-2026">July 2026</option>
                    <option value="-08-2026">August 2026</option>
                  </>
                )}
              </select>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <span style={{ fontSize: '0.65rem', color: '#64748b', fontWeight: 'bold' }}>Lead Stage</span>
              <select
                value={filters.stage}
                onChange={(e) => setFilters({ ...filters, stage: e.target.value })}
                style={{ border: '1px solid #cbd5e1', borderRadius: '4px', padding: '3px 6px', fontSize: '0.76rem', backgroundColor: '#fff', color: '#334155' }}
              >
                <option value="">Select Here</option>
                {getUniqueOptions(/stage|status/i).length > 0 ? (
                  getUniqueOptions(/stage|status/i).map((s) => <option key={s} value={s}>{s}</option>)
                ) : (
                  <>
                    <option value="Qualified">Qualified</option>
                    <option value="Contacted">Contacted</option>
                    <option value="New Lead">New Lead</option>
                    <option value="Fee Paid">Fee Paid</option>
                  </>
                )}
              </select>
            </div>

            {isAdmin && (
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <span style={{ fontSize: '0.65rem', color: '#64748b', fontWeight: 'bold' }}>Counselor / Owner</span>
                <select
                  value={filters.owner}
                  onChange={(e) => setFilters({ ...filters, owner: e.target.value })}
                  style={{ border: '1px solid #cbd5e1', borderRadius: '4px', padding: '3px 6px', fontSize: '0.76rem', backgroundColor: '#fff', color: '#334155' }}
                >
                  <option value="">All Counselors</option>
                  {getUniqueOptions(/owner|counselor|team/i).length > 0 ? (
                    getUniqueOptions(/owner|counselor|team/i).map((o) => <option key={o} value={o}>{o}</option>)
                  ) : (
                    <>
                      <option value="Counselor Jane Smith">Jane Smith</option>
                      <option value="Counselor Priya Sharma">Priya Sharma</option>
                      <option value="Counselor Alex Rivera">Alex Rivera</option>
                    </>
                  )}
                </select>
              </div>
            )}

            <div style={{ display: 'flex', flexDirection: 'column' }}>
              <span style={{ fontSize: '0.65rem', color: '#64748b', fontWeight: 'bold' }}>Campaign Source</span>
              <select
                value={filters.campaign}
                onChange={(e) => setFilters({ ...filters, campaign: e.target.value })}
                style={{ border: '1px solid #cbd5e1', borderRadius: '4px', padding: '3px 6px', fontSize: '0.76rem', backgroundColor: '#fff', color: '#334155' }}
              >
                <option value="">Select Here</option>
                {getUniqueOptions(/campaign|source/i).length > 0 ? (
                  getUniqueOptions(/campaign|source/i).map((c) => <option key={c} value={c}>{c}</option>)
                ) : (
                  <>
                    <option value="Social Media">Social Media</option>
                    <option value="Google Ads">Google Ads</option>
                    <option value="Meta Ads">Meta Ads</option>
                    <option value="Email Marketing">Email Marketing</option>
                    <option value="Organic Search">Organic Search</option>
                    <option value="Website">Website</option>
                  </>
                )}
              </select>
            </div>

            <button
              onClick={() => setMoreFiltersModalOpen(true)}
              style={{ background: 'none', border: 'none', color: '#2563eb', fontSize: '0.75rem', fontWeight: '700', cursor: 'pointer', marginTop: '10px' }}
            >
              +5 more
            </button>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <button
              onClick={handleResetFilters}
              style={{ background: 'none', border: 'none', color: '#64748b', fontSize: '0.78rem', cursor: 'pointer' }}
            >
              Reset
            </button>
            <span style={{ fontSize: '0.76rem', color: '#64748b' }}>
              Showing <strong>{Math.max(currentSheetRows.length - 1, 0)}</strong> leads
            </span>
          </div>
        </div>

        {/* 2D CANVAS VIEW */}
        <div style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column', padding: '14px 20px' }}>
          {workbookData ? (
            <div style={{ backgroundColor: '#ffffff', borderRadius: '8px', border: '1px solid #cbd5e1', flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              
              {/* Formula Bar + Batch Delete Rows */}
              <div style={{ padding: '6px 12px', backgroundColor: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px', fontSize: '0.8rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1 }}>
                  <strong style={{ color: '#2563eb', minWidth: '40px' }}>
                    {getColLetter(activeCell[1])}{activeCell[0] + 1}
                  </strong>
                  <span style={{ color: '#94a3b8', fontWeight: 'bold' }}>fx</span>
                  <div style={{ flex: 1, backgroundColor: '#fff', padding: '3px 8px', borderRadius: '4px', border: '1px solid #cbd5e1', minHeight: '18px', color: '#334155' }}>
                    {String(activeCellValue)}
                  </div>
                </div>

                {isAdmin && selectedRowIndices.length > 0 && (
                  <button
                    onClick={handleDeleteCheckedRows}
                    style={{
                      backgroundColor: '#dc2626',
                      color: '#ffffff',
                      border: 'none',
                      padding: '5px 12px',
                      borderRadius: '5px',
                      fontSize: '0.78rem',
                      cursor: 'pointer',
                      fontWeight: '700',
                    }}
                  >
                    Delete Selected ({selectedRowIndices.length})
                  </button>
                )}
              </div>

              {/* 2D Excel Grid */}
              <div style={{ flex: 1, overflow: 'auto' }}>
                <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: '0.82rem', tableLayout: 'fixed' }}>
                  <thead>
                    <tr style={{ backgroundColor: '#f1f5f9', position: 'sticky', top: 0, zIndex: 2 }}>
                      {isAdmin && (
                        <th style={{ width: '40px', border: '1px solid #cbd5e1', padding: '6px', textAlign: 'center' }}>
                          <input
                            type="checkbox"
                            checked={selectedRowIndices.length === currentSheetRows.length - 1 && currentSheetRows.length > 1}
                            onChange={handleToggleSelectAll}
                            title="Select all rows"
                          />
                        </th>
                      )}
                      <th style={{ width: '50px', border: '1px solid #cbd5e1', padding: '6px', textAlign: 'center', color: '#64748b' }}>#</th>
                      {Array.from({ length: totalCols }).map((_, cIdx) => (
                        <th key={cIdx} style={{ width: '160px', border: '1px solid #cbd5e1', padding: '6px 8px', textAlign: 'center', color: '#475569', fontWeight: 'bold' }}>
                          {getColLetter(cIdx)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentSheetRows.map((row, rIdx) => {
                      const isChecked = selectedRowIndices.includes(rIdx);
                      return (
                        <tr key={rIdx} style={{ backgroundColor: isChecked ? '#fef2f2' : 'transparent' }}>
                          {isAdmin && (
                            <td style={{ border: '1px solid #cbd5e1', textAlign: 'center', padding: '4px' }}>
                              {rIdx > 0 ? (
                                <input
                                  type="checkbox"
                                  checked={isChecked}
                                  onChange={() => handleToggleRow(rIdx)}
                                />
                              ) : null}
                            </td>
                          )}
                          <td style={{ border: '1px solid #cbd5e1', backgroundColor: '#f8fafc', textAlign: 'center', color: '#64748b', fontWeight: '600', padding: '4px' }}>
                            {rIdx + 1}
                          </td>
                          {Array.from({ length: totalCols }).map((_, cIdx) => {
                            const val = row[cIdx] !== undefined ? row[cIdx] : '';
                            const isSelected = activeCell[0] === rIdx && activeCell[1] === cIdx;
                            return (
                              <td
                                key={cIdx}
                                onClick={() => setActiveCell([rIdx, cIdx])}
                                style={{
                                  border: '1px solid #e2e8f0',
                                  padding: '6px 10px',
                                  color: '#0f172a',
                                  whiteSpace: 'nowrap',
                                  overflow: 'hidden',
                                  textOverflow: 'ellipsis',
                                  backgroundColor: isSelected ? '#eff6ff' : rIdx === 0 ? '#f8fafc' : isChecked ? '#fef2f2' : '#ffffff',
                                  outline: isSelected ? '2px solid #2563eb' : 'none',
                                  outlineOffset: '-2px',
                                  fontWeight: rIdx === 0 ? '600' : 'normal',
                                  cursor: 'pointer',
                                }}
                              >
                                {String(val ?? '')}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Bottom Sheet Switcher Tabs */}
              <div style={{ backgroundColor: '#f1f5f9', borderTop: '1px solid #cbd5e1', padding: '5px 12px', display: 'flex', gap: '6px', alignItems: 'center' }}>
                <span style={{ fontSize: '0.72rem', fontWeight: 'bold', color: '#64748b', marginRight: '6px' }}>WORKBOOK TABS:</span>
                {workbookData.sheetNames.map((name) => (
                  <button
                    key={name}
                    onClick={() => {
                      setActiveSheetName(name);
                      setActiveCell([0, 0]);
                      setSelectedRowIndices([]);
                    }}
                    style={{
                      padding: '4px 12px',
                      borderRadius: '4px 4px 0 0',
                      border: '1px solid #cbd5e1',
                      borderBottom: activeSheetName === name ? '2px solid #2563eb' : '1px solid #cbd5e1',
                      backgroundColor: activeSheetName === name ? '#ffffff' : '#e2e8f0',
                      color: activeSheetName === name ? '#2563eb' : '#475569',
                      fontWeight: activeSheetName === name ? 'bold' : 'normal',
                      fontSize: '0.78rem',
                      cursor: 'pointer',
                    }}
                  >
                    {name}
                  </button>
                ))}
              </div>

            </div>
          ) : (
            <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ backgroundColor: '#ffffff', borderRadius: '12px', border: '2px dashed #cbd5e1', padding: '50px 40px', textAlign: 'center', maxWidth: '480px', width: '100%' }}>
                <h3 style={{ margin: '0 0 6px 0', fontSize: '1.2rem', color: '#0f172a' }}>
                  {isAdmin ? 'No Excel File Imported' : 'No Active Leads File Found'}
                </h3>
                <p style={{ margin: '0 0 16px 0', fontSize: '0.85rem', color: '#64748b' }}>
                  {isAdmin
                    ? 'Import any .xlsx or .csv spreadsheet to activate the Leads Management System workspace.'
                    : 'Your administrator has not loaded a leads spreadsheet into the workspace yet.'}
                </p>
                {isAdmin && (
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    style={{ backgroundColor: '#2563eb', color: '#ffffff', border: 'none', padding: '9px 20px', borderRadius: '6px', fontWeight: '600', fontSize: '0.85rem', cursor: 'pointer' }}
                  >
                    Select Excel File
                  </button>
                )}
              </div>
            </div>
          )}
        </div>
      </main>

      {/* SETTINGS MODAL (ADMIN ONLY) */}
      {settingsModalOpen && isAdmin && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#ffffff', borderRadius: '8px', width: '420px', padding: '22px', boxShadow: '0 10px 25px rgba(0,0,0,0.15)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h4 style={{ margin: 0, color: '#0f172a', fontSize: '1.05rem' }}>Administrator Settings</h4>
              <button onClick={() => setSettingsModalOpen(false)} style={{ background: 'none', border: 'none', fontSize: '1.1rem', cursor: 'pointer', color: '#64748b' }}>✕</button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', fontSize: '0.85rem' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={autoTrimOnImport}
                  onChange={(e) => setAutoTrimOnImport(e.target.checked)}
                />
                <span>Auto-trim leading/trailing cell whitespace on import</span>
              </label>

              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={confirmBeforeDelete}
                  onChange={(e) => setConfirmBeforeDelete(e.target.checked)}
                />
                <span>Show confirmation dialog before batch deleting rows</span>
              </label>

              <div style={{ borderTop: '1px solid #e2e8f0', paddingTop: '12px', marginTop: '6px' }}>
                <span style={{ display: 'block', fontSize: '0.78rem', color: '#64748b', marginBottom: '8px' }}>Workspace Cache:</span>
                <button
                  onClick={() => {
                    setWorkbookData(null);
                    setActiveSheetName('');
                    setActiveCell([0, 0]);
                    setSelectedRowIndices([]);
                    handleResetFilters();
                    setSettingsModalOpen(false);
                    alert('Workspace cache cleared.');
                  }}
                  style={{ backgroundColor: '#fef2f2', color: '#dc2626', border: '1px solid #fecaca', padding: '6px 12px', borderRadius: '5px', fontSize: '0.8rem', fontWeight: '600', cursor: 'pointer' }}
                >
                  Clear Loaded Spreadsheet
                </button>
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '20px' }}>
              <button
                onClick={() => setSettingsModalOpen(false)}
                style={{ padding: '6px 16px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '5px', fontSize: '0.82rem', fontWeight: 'bold', cursor: 'pointer' }}
              >
                Save & Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* +5 MORE FILTERS MODAL */}
      {moreFiltersModalOpen && (
        <div style={{ position: 'fixed', inset: 0, backgroundColor: 'rgba(15, 23, 42, 0.5)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 100 }}>
          <div style={{ backgroundColor: '#ffffff', borderRadius: '8px', width: '400px', padding: '20px', boxShadow: '0 10px 25px rgba(0,0,0,0.15)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
              <h4 style={{ margin: 0, color: '#0f172a', fontSize: '1rem' }}>Additional Lead Field Filters</h4>
              <button onClick={() => setMoreFiltersModalOpen(false)} style={{ background: 'none', border: 'none', fontSize: '1.1rem', cursor: 'pointer', color: '#64748b' }}>✕</button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 'bold', color: '#475569', marginBottom: '4px' }}>Choose Lead Column:</label>
                <select
                  value={selectedCustomCol}
                  onChange={(e) => setSelectedCustomCol(e.target.value)}
                  style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '0.82rem' }}
                >
                  <option value="">Select column...</option>
                  {(workbookData?.sheets[activeSheetName]?.[0] || []).map((header, idx) => (
                    <option key={idx} value={idx}>{header || `Column ${getColLetter(idx)}`}</option>
                  ))}
                </select>
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: 'bold', color: '#475569', marginBottom: '4px' }}>Filter Value Contains:</label>
                <input
                  type="text"
                  placeholder="e.g. B.Tech, Verified, Scholarship..."
                  value={customColValue}
                  onChange={(e) => setCustomColValue(e.target.value)}
                  style={{ width: '100%', padding: '6px 8px', border: '1px solid #cbd5e1', borderRadius: '5px', fontSize: '0.82rem', boxSizing: 'border-box' }}
                />
              </div>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '18px' }}>
              <button
                onClick={() => { setSelectedCustomCol(''); setCustomColValue(''); setMoreFiltersModalOpen(false); }}
                style={{ padding: '6px 12px', border: '1px solid #cbd5e1', borderRadius: '5px', background: '#fff', fontSize: '0.8rem', cursor: 'pointer' }}
              >
                Clear & Close
              </button>
              <button
                onClick={() => setMoreFiltersModalOpen(false)}
                style={{ padding: '6px 14px', backgroundColor: '#2563eb', color: '#fff', border: 'none', borderRadius: '5px', fontSize: '0.8rem', fontWeight: 'bold', cursor: 'pointer' }}
              >
                Apply Lead Filter
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}