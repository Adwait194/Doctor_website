import { useEffect, useState } from 'react'
import * as XLSX from 'xlsx'
import { supabase } from './supabase'

// ─── Nested options data ─────────────────────────────────────────────────────

const diseaseOptions = [
  {
    label: 'गोवर',
    sub: ['अतिसार', 'नियमीत', '१ दिवसाआड', 'बद्धकोष्ठता'],
  },
  {
    label: 'कांजिण्या',
    sub: ['वारंवार', 'जळजळ', 'नक्त'],
  },
  {
    label: 'कावीळ',
    sub: ['चांगला', 'मंद', 'अनियमित', 'अतिशय मंद', 'विकृत'],
  },
  {
    label: 'Diabetes',
    sub: ['चांगला', 'मंद', 'अनियमित', 'अतिशय मंद', 'विकृत'],
  },
  {
    label: 'Hyper Tension',
    sub: ['चांगला', 'मंद', 'अनियमित', 'अतिशय मंद', 'विकृत'],
  },
  
]

const symptomOptions = [
  {
    label: 'मल',
    sub: ['सामान्य', 'सैल / अतिसार', 'कठीण / बद्धकोष्ठता','नियमीत',' १ दिवसाआड'],
  },
  {
    label: 'मूत्र',
    sub: ['सामान्य', 'वारंवार', 'जळजळ','नक्त'],
  },
  {
    label: 'अग्नि',
    sub: ['चांगला', 'मंद', 'अनियमित', 'अतिशय मंद', 'विकृत'],
  },
  {
    label: 'निद्रा',
    sub: ['चांगली', 'कमी', 'खंडित', 'जास्त'],
  },
  {
    label: 'उदरपरिक्षण',
    sub: ['यकृत पीडनासहत्व','यकृत घन','यकृत आकाशीय','प्लिहा पीडनासहत्व','प्लिहा घन','प्लिहा आकाशीय','द .वृक्क','वामवृक्क',' पीडनासहत्व वृक्क', 'द .गर्भाशय',' वाम गर्भाशय','पीडनासहत्व गर्भाशय',],
  },
  {
    label: 'नाडी',
    sub: ['साम', 'कठिण', 'स्थूल','तीक्ष्ण','मंद','जलद'],
  },
  {
    label: 'प्रकृति',
    sub: ['अम्ल','क्षार'],
  },
] 

// ─── Nested Checkbox Component ────────────────────────────────────────────────

function NestedCheckbox({ option, selected, onChange, accentColor }) {
  const [open, setOpen] = useState(false)

  const parentKey = option.label
  const isParentChecked = selected.includes(parentKey)
  const checkedSubs = option.sub.filter((s) => selected.includes(`${parentKey}::${s}`))
  const someSubChecked = checkedSubs.length > 0

  function toggleParent() {
    if (isParentChecked || someSubChecked) {
      // remove parent + all subs
      const remove = [parentKey, ...option.sub.map((s) => `${parentKey}::${s}`)]
      onChange(selected.filter((x) => !remove.includes(x)))
    } else {
      onChange([...selected, parentKey])
    }
  }

  function toggleSub(sub) {
    const key = `${parentKey}::${sub}`
    if (selected.includes(key)) {
      const next = selected.filter((x) => x !== key)
      // if no subs remain, also remove parent
      const remainingSubs = option.sub.filter((s) => next.includes(`${parentKey}::${s}`))
      if (remainingSubs.length === 0) {
        onChange(next.filter((x) => x !== parentKey))
      } else {
        onChange(next)
      }
    } else {
      // add sub and ensure parent is in list
      const next = selected.filter((x) => x !== parentKey)
      onChange([...next, parentKey, key])
    }
  }

  return (
    <div style={styles.nestedItem}>
      <div style={styles.parentRow}>
        <label style={styles.checkLabel}>
          <div
            style={{
              ...styles.customCheck,
              background: isParentChecked || someSubChecked ? accentColor : 'transparent',
              borderColor: isParentChecked || someSubChecked ? accentColor : '#c9b99a',
            }}
            onClick={toggleParent}
          >
            {(isParentChecked || someSubChecked) && (
              <svg width="10" height="8" viewBox="0 0 10 8" fill="none">
                <path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            )}
          </div>
          <span style={styles.parentLabel} onClick={toggleParent}>
            {option.label}
          </span>
        </label>

        <button
          type="button"
          style={{ ...styles.expandBtn, color: accentColor }}
          onClick={() => setOpen((o) => !o)}
        >
          {open ? '▲' : '▼'}
        </button>
      </div>

      {open && (
        <div style={styles.subList}>
          {option.sub.map((sub) => {
            const key = `${parentKey}::${sub}`
            const checked = selected.includes(key)
            return (
              <label key={sub} style={styles.subLabel} onClick={() => toggleSub(sub)}>
                <div
                  style={{
                    ...styles.customCheck,
                    width: 14,
                    height: 14,
                    background: checked ? accentColor : 'transparent',
                    borderColor: checked ? accentColor : '#c9b99a',
                  }}
                >
                  {checked && (
                    <svg width="8" height="6" viewBox="0 0 10 8" fill="none">
                      <path d="M1 4L3.5 6.5L9 1" stroke="white" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  )}
                </div>
                <span style={{ fontSize: 13, color: '#5a4a3a' }}>{sub}</span>
              </label>
            )
          })}
        </div>
      )}

      {someSubChecked && (
        <div style={styles.subTags}>
          {checkedSubs.map((s) => (
            <span key={s} style={{ ...styles.tag, background: accentColor + '22', color: accentColor, border: `1px solid ${accentColor}55` }}>
              {s}
            </span>
          ))}
        </div>
      )}
    </div>
  )
}

// ─── Section Card ─────────────────────────────────────────────────────────────

function Card({ title, icon, children, accent = '#2e6b5e' }) {
  return (
    <div style={styles.card}>
      <div style={{ ...styles.cardHeader, borderLeft: `4px solid ${accent}` }} className="card-header">
        <span style={{ fontSize: 20, marginRight: 10 }}>{icon}</span>
        <h2 style={{ ...styles.cardTitle, color: accent }} className="card-title">{title}</h2>
      </div>
      <div style={styles.cardBody} className="card-body">{children}</div>
    </div>
  )
}

// ─── Input / Select helpers ───────────────────────────────────────────────────

function Field({ label, children }) {
  return (
    <div style={styles.field}>
      <label style={styles.fieldLabel}>{label}</label>
      {children}
    </div>
  )
}

const inp = {
  width: '100%',
  padding: '10px 14px',
  border: '1.5px solid #d9cbb8',
  borderRadius: 8,
  fontSize: 14,
  color: '#3a2e22',
  background: '#fffdf8',
  outline: 'none',
  boxSizing: 'border-box',
  fontFamily: "'DM Sans', sans-serif",
}

// ─── Age calculation helper ────────────────────────────────────────────────────

function calculateAge(birthdate) {
  if (!birthdate) return ''
  const dob = new Date(birthdate)
  if (isNaN(dob.getTime())) return ''
  const today = new Date()
  let age = today.getFullYear() - dob.getFullYear()
  const m = today.getMonth() - dob.getMonth()
  if (m < 0 || (m === 0 && today.getDate() < dob.getDate())) {
    age--
  }
  return age >= 0 ? age : ''
}

// ─── Tag Input (for multiple page numbers) ─────────────────────────────────────

function TagInput({ values, onChange, placeholder, accentColor = '#2e6b5e' }) {
  const [draft, setDraft] = useState('')

  function commitDraft() {
    const v = draft.trim()
    if (v && !values.includes(v)) {
      onChange([...values, v])
    }
    setDraft('')
  }

  function handleKeyDown(e) {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault()
      commitDraft()
    } else if (e.key === 'Backspace' && !draft && values.length > 0) {
      onChange(values.slice(0, -1))
    }
  }

  function removeTag(v) {
    onChange(values.filter((x) => x !== v))
  }

  return (
    <div>
      <div style={{ ...inp, display: 'flex', flexWrap: 'wrap', gap: 6, alignItems: 'center', minHeight: 42, padding: '6px 10px' }}>
        {values.map((v) => (
          <span key={v} style={{ ...styles.tag, background: accentColor + '22', color: accentColor, border: `1px solid ${accentColor}55`, display: 'flex', alignItems: 'center', gap: 6 }}>
            {v}
            <span onClick={() => removeTag(v)} style={{ cursor: 'pointer', fontWeight: 700, lineHeight: 1 }}>×</span>
          </span>
        ))}
        <input
          type="text"
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={commitDraft}
          placeholder={values.length === 0 ? placeholder : ''}
          style={{ border: 'none', outline: 'none', flex: 1, minWidth: 80, fontSize: 14, fontFamily: "'DM Sans', sans-serif", background: 'transparent', color: '#3a2e22' }}
        />
      </div>
    </div>
  )
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [patients, setPatients] = useState([])
  const [visits, setVisits] = useState([])
  const [selectedPatient, setSelectedPatient] = useState(null)
  const [selectedPatientData, setSelectedPatientData] = useState(null)
  const [searchPhone, setSearchPhone] = useState('')
  const [editingPatientId, setEditingPatientId] = useState(null)
  const [activeTab, setActiveTab] = useState('patients') // 'patients' | 'visits'

  const emptyForm = { name: '', phone: '', gender: '', birthdate: '', age: '', weight: '', page_no: [], past_disease: [] }
  const emptyVisit = { symptoms: [], diagnosis: '', medicines: '', doctor_notes: '', fee: '' }

  const [form, setForm] = useState(emptyForm)
  const [visitForm, setVisitForm] = useState(emptyVisit)

  useEffect(() => { fetchPatients() }, [])

  async function fetchPatients() {
    const { data, error } = await supabase.from('patients').select('*').order('patient_id', { ascending: false })
    if (!error) setPatients(data)
  }

  async function addPatient(e) {
    e.preventDefault()
    const { error } = await supabase.from('patients').insert([form])
    if (error) { alert(error.message) } else {
      alert('Patient Added ✓')
      setForm(emptyForm)
      fetchPatients()
    }
  }

  async function updatePatient() {
    const { error } = await supabase.from('patients').update(form).eq('patient_id', editingPatientId)
    if (error) { alert(error.message) } else {
      alert('Patient Updated ✓')
      setEditingPatientId(null)
      setForm(emptyForm)
      fetchPatients()
    }
  }
async function searchPatient() {
  const query = searchPhone.trim()
  if (!query) {
    fetchPatients()
    return
  }
  const { data, error } = await supabase
    .from('patients')
    .select('*')
    .or(`phone.ilike.%${query}%,name.ilike.%${query}%`)
  if (!error) setPatients(data)
}

  async function loadVisits(patient) {
    setSelectedPatient(patient.patient_id)
    setSelectedPatientData(patient)
    setActiveTab('visits')
    const { data, error } = await supabase.from('visits').select('*').eq('patient_id', patient.patient_id).order('visit_date', { ascending: false })
    if (!error) setVisits(data)
  }

  async function addVisit(e) {
    e.preventDefault()
    const { error } = await supabase.from('visits').insert([{ patient_id: selectedPatient, ...visitForm }])
    if (error) { alert(error.message) } else {
      alert('Visit Added ✓')
      setVisitForm(emptyVisit)
      loadVisits(selectedPatientData)
    }
  }

  // Flatten selected values to store (keep parent key if no sub selected, else only sub keys)
  function formatSelectedForDisplay(arr) {
    return arr.map((x) => x.replace('::', ' → ')).join(', ')
  }

  function exportPatientsToExcel() {
    if (patients.length === 0) { alert('No patients to export'); return }
    const rows = patients.map((p) => ({
      ID: p.patient_id,
      Name: p.name,
      Phone: p.phone,
      Gender: p.gender,
      'Date of Birth': p.birthdate,
      'Age (yrs)': p.age,
      'Weight (kg)': p.weight,
      'Page No.': (p.page_no || []).join(', '),
      'Past Diseases': formatSelectedForDisplay(p.past_disease || []),
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Patients')
    XLSX.writeFile(wb, `patients_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  async function exportVisitsToExcel() {
    // Export visits for the currently selected patient, or all visits if none selected
    let data = visits
    let filename = 'visits'
    if (selectedPatient) {
      filename = `visits_${selectedPatientData?.name || selectedPatient}`
    } else {
      const { data: allVisits, error } = await supabase.from('visits').select('*').order('visit_date', { ascending: false })
      if (error) { alert(error.message); return }
      data = allVisits
    }
    if (!data || data.length === 0) { alert('No visits to export'); return }
    const rows = data.map((v) => ({
      'Visit ID': v.visit_id,
      'Patient ID': v.patient_id,
      'Visit Date': v.visit_date ? new Date(v.visit_date).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' }) : '',
      Symptoms: formatSelectedForDisplay(v.symptoms || []),
      Diagnosis: v.diagnosis,
      Medicines: v.medicines,
      'Fee Taken (₹)': v.fee,
      'Doctor Notes': v.doctor_notes,
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Visits')
    XLSX.writeFile(wb, `${filename}_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;700&family=DM+Sans:wght@300;400;500&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { background: #f5f0e8; font-family: 'DM Sans', sans-serif; color: #3a2e22; }
        input:focus, select:focus, textarea:focus { border-color: #2e6b5e !important; box-shadow: 0 0 0 3px #2e6b5e22; }
        ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-thumb { background: #c9b99a; border-radius: 3px; }
        tr:hover td { background: #f9f5ed; }

        .table-scroll { -webkit-overflow-scrolling: touch; }
        .tabs-scroll { -webkit-overflow-scrolling: touch; }

        @media (max-width: 768px) {
          .app-header-inner { flex-direction: column; align-items: flex-start; gap: 12px; padding: 16px 18px !important; }
          .app-header-stats { align-self: stretch; }
          .app-main { padding: 0 12px 40px !important; }
          .two-col { grid-template-columns: 1fr !important; gap: 12px !important; }
          .card-body { padding: 16px 14px !important; }
          .card-header { padding: 14px 16px 10px !important; }
          .card-title { font-size: 16px !important; }
          .tabs-scroll { overflow-x: auto; }
          .tab-btn { padding: 10px 16px !important; font-size: 13px !important; white-space: nowrap; }
          .search-row { flex-direction: column !important; }
          .search-row > button { width: 100%; }
          .export-row { justify-content: stretch !important; }
          .export-row > button { width: 100%; justify-content: center; }
          .form-actions { flex-direction: column !important; }
          .form-actions > button { width: 100%; }
          .patient-banner { flex-direction: column; align-items: flex-start !important; gap: 10px !important; padding: 14px 16px !important; }
          .visit-grid { grid-template-columns: 1fr !important; }
          .table-min-width { min-width: 640px; }
        }

        @media (max-width: 480px) {
          .logo-title { font-size: 18px !important; }
          .logo-icon { font-size: 32px !important; }
          .stat-num { font-size: 20px !important; }
          .header-stat { padding: 6px 14px !important; }
        }
      `}</style>

      {/* ── Header ── */}
      <header style={styles.header}>
        <div style={styles.headerInner} className="app-header-inner">
          <div style={styles.logo}>
            <span style={styles.logoIcon} className="logo-icon">⚕</span>
            <div>
              <div style={styles.logoTitle} className="logo-title">Clinic Management</div>
              <div style={styles.logoSub}>आयुर्वेद दवाखाना • Patient Records</div>
            </div>
          </div>
          <div style={styles.headerStats} className="app-header-stats">
            <div style={styles.stat} className="header-stat"><span style={styles.statNum} className="stat-num">{patients.length}</span><span style={styles.statLbl}>Patients</span></div>
          </div>
        </div>
      </header>

      <div style={styles.main} className="app-main">

        {/* ── Tabs ── */}
        <div style={styles.tabs} className="tabs-scroll">
          {['patients', 'visits'].map((t) => (
            <button key={t} type="button" onClick={() => setActiveTab(t)}
              className="tab-btn"
              style={{ ...styles.tab, ...(activeTab === t ? styles.tabActive : {}) }}>
              {t === 'patients' ? '🧑‍⚕️ Patients' : '📋 Visits'}
              {t === 'visits' && selectedPatientData && (
                <span style={styles.tabBadge}>{selectedPatientData.name}</span>
              )}
            </button>
          ))}
        </div>

        {/* ══════════ PATIENTS TAB ══════════ */}
        {activeTab === 'patients' && (
          <div style={styles.grid}>

            {/* Search */}
            <Card title="Search Patient" icon="🔍" accent="#7b5ea7">
              <div style={{ display: 'flex', gap: 10 }} className="search-row">
                <input style={{ ...inp, flex: 1 }} type="text" placeholder="Enter phone number or Name"
                  value={searchPhone} onChange={(e) => setSearchPhone(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && searchPatient()} />
                <button type="button" style={styles.btnSecondary} onClick={searchPatient}>Search</button>
                <button type="button" style={{ ...styles.btnSecondary, background: '#f5f0e8', color: '#7b5ea7' }}
                  onClick={() => { setSearchPhone(''); fetchPatients() }}>All</button>
              </div>
            </Card>

            {/* Add / Edit Patient */}
            <Card title={editingPatientId ? 'Update Patient' : 'Add New Patient'} icon="➕" accent="#2e6b5e">
              <form onSubmit={editingPatientId ? (e) => { e.preventDefault(); updatePatient() } : addPatient}>
                <div style={styles.twoCol} className="two-col">
                  <Field label="Full Name">
                    <input style={inp} type="text" placeholder="Patient name" value={form.name}
                      onChange={(e) => setForm({ ...form, name: e.target.value })} required />
                  </Field>
                  <Field label="Phone">
                    <input style={inp} type="text" placeholder="Mobile number" value={form.phone}
                      onChange={(e) => setForm({ ...form, phone: e.target.value })} />
                  </Field>
                </div>

                <div style={styles.twoCol} className="two-col">
                  <Field label="Gender">
                    <select style={inp} value={form.gender} onChange={(e) => setForm({ ...form, gender: e.target.value })}>
                      <option value="">Select Gender</option>
                      <option>Male</option>
                      <option>Female</option>
                      <option>Other</option>
                    </select>
                  </Field>
                  <Field label="Date of Birth">
                    <input style={inp} type="date" value={form.birthdate}
                      onChange={(e) => {
                        const birthdate = e.target.value
                        setForm({ ...form, birthdate, age: calculateAge(birthdate) })
                      }} />
                  </Field>
                </div>

                <div style={styles.twoCol} className="two-col">
                  <Field label="Age (years)">
                    <input style={inp} type="number" placeholder="Age" value={form.age}
                      onChange={(e) => setForm({ ...form, age: e.target.value })} />
                  </Field>
                  <Field label="Weight (kg)">
                    <input style={inp} type="number" placeholder="Weight" value={form.weight}
                      onChange={(e) => setForm({ ...form, weight: e.target.value })} />
                  </Field>
                </div>

                <Field label="Page No.">
                  <TagInput
                    values={form.page_no}
                    onChange={(val) => setForm({ ...form, page_no: val })}
                    placeholder="Type a page number and press Enter…"
                    accentColor="#2e6b5e"
                  />
                </Field>

                <div style={styles.checkSection}>
                  <div style={styles.checkTitle}>📌 भूतपूर्व रोग (Past Diseases)</div>
                  {diseaseOptions.map((opt) => (
                    <NestedCheckbox key={opt.label} option={opt} accentColor="#2e6b5e"
                      selected={form.past_disease}
                      onChange={(val) => setForm({ ...form, past_disease: val })} />
                  ))}
                </div>

                <div style={{ display: 'flex', gap: 10, marginTop: 8 }} className="form-actions">
                  <button type="submit" style={styles.btnPrimary}>
                    {editingPatientId ? '✔ Update Patient' : '+ Add Patient'}
                  </button>
                  {editingPatientId && (
                    <button type="button" style={styles.btnGhost}
                      onClick={() => { setEditingPatientId(null); setForm(emptyForm) }}>
                      Cancel
                    </button>
                  )}
                </div>
              </form>
            </Card>

            {/* Patients Table */}
            <Card title={`Patients (${patients.length})`} icon="👥" accent="#c47c2b">
              <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }} className="export-row">
                <button type="button" style={{ ...styles.btnSecondary, background: '#2e6b5e' }} onClick={exportPatientsToExcel}>
                  📥 Export to Excel
                </button>
              </div>
              <div style={{ overflowX: 'auto' }} className="table-scroll">
                <table style={styles.table} className="table-min-width">
                  <thead>
                    <tr>
                      {['ID', 'Name', 'Phone', 'Gender', 'Age', 'Weight', 'Page No.', 'Actions'].map((h) => (
                        <th key={h} style={styles.th}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {patients.length === 0 && (
                      <tr><td colSpan={8} style={{ textAlign: 'center', padding: 30, color: '#9a8a7a' }}>No patients found</td></tr>
                    )}
                    {patients.map((p) => (
                      <tr key={p.patient_id}>
                        <td style={styles.td}>
                          <span style={styles.idBadge}>#{p.patient_id}</span>
                        </td>
                        <td style={{ ...styles.td, fontWeight: 500 }}>{p.name}</td>
                        <td style={styles.td}>{p.phone}</td>
                        <td style={styles.td}>
                          <span style={{ ...styles.genderBadge, background: p.gender === 'Female' ? '#f9d8f0' : p.gender === 'Male' ? '#d8eaf9' : '#f0f0f0' }}>
                            {p.gender}
                          </span>
                        </td>
                        <td style={styles.td}>{p.age} yr</td>
                        <td style={styles.td}>{p.weight} kg</td>
                        <td style={styles.td}>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                            {(p.page_no || []).map((pn) => (
                              <span key={pn} style={{ ...styles.tag, background: '#f0e8d8', color: '#7a6a5a', border: '1px solid #e0d5c5' }}>{pn}</span>
                            ))}
                          </div>
                        </td>
                        <td style={styles.td}>
                          <div style={{ display: 'flex', gap: 6 }}>
                            <button style={styles.actionBtn} onClick={() => {
                              setForm({ name: p.name, phone: p.phone, gender: p.gender, birthdate: p.birthdate, age: p.birthdate ? calculateAge(p.birthdate) : p.age, weight: p.weight, page_no: p.page_no || [], past_disease: p.past_disease || [] })
                              setEditingPatientId(p.patient_id)
                              window.scrollTo({ top: 0, behavior: 'smooth' })
                            }}>✏ Edit</button>
                            <button style={{ ...styles.actionBtn, background: '#2e6b5e', color: '#fff' }}
                              onClick={() => loadVisits(p)}>📋 Visits</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          </div>
        )}

        {/* ══════════ VISITS TAB ══════════ */}
        {activeTab === 'visits' && (
          <div style={styles.grid}>
            {!selectedPatient ? (
              <div style={styles.emptyState}>
                <div style={{ fontSize: 48 }}>📋</div>
                <div style={{ color: '#9a8a7a', marginTop: 12 }}>Select a patient from the Patients tab to manage visits.</div>
              </div>
            ) : (
              <>
                {/* Patient Info Banner */}
                <div style={styles.patientBanner} className="patient-banner">
                  <div style={{ fontSize: 36 }}>🧑</div>
                  <div>
                    <div style={{ fontSize: 18, fontWeight: 600 }}>{selectedPatientData?.name}</div>
                    <div style={{ fontSize: 13, color: '#7a8a7a' }}>{selectedPatientData?.phone} • {selectedPatientData?.gender} • {selectedPatientData?.age} yrs • {selectedPatientData?.weight} kg</div>
                  </div>
                </div>

                {/* Add Visit */}
                <Card title="Record New Visit" icon="🩺" accent="#2e6b5e">
                  <form onSubmit={addVisit}>
                    <div style={styles.checkSection}>
                      <div style={styles.checkTitle}>🔬 लक्षणे (Symptoms)</div>
                      {symptomOptions.map((opt) => (
                        <NestedCheckbox key={opt.label} option={opt} accentColor="#c47c2b"
                          selected={visitForm.symptoms}
                          onChange={(val) => setVisitForm({ ...visitForm, symptoms: val })} />
                      ))}
                    </div>

                    <div style={styles.twoCol} className="two-col">
                      <Field label="Diagnosis / निदान">
                        <textarea style={{ ...inp, height: 90, resize: 'vertical' }}
                          placeholder="Enter diagnosis…"
                          value={visitForm.diagnosis}
                          onChange={(e) => setVisitForm({ ...visitForm, diagnosis: e.target.value })} />
                      </Field>
                      <Field label="Medicines / औषधे">
                        <textarea style={{ ...inp, height: 90, resize: 'vertical' }}
                          placeholder="Prescribed medicines…"
                          value={visitForm.medicines}
                          onChange={(e) => setVisitForm({ ...visitForm, medicines: e.target.value })} />
                      </Field>
                    </div>

                    <div style={styles.twoCol} className="two-col">
                      <Field label="Fee Taken (₹)">
                        <input style={inp} type="number" placeholder="Fee amount" value={visitForm.fee}
                          onChange={(e) => setVisitForm({ ...visitForm, fee: e.target.value })} />
                      </Field>
                    </div>

                    <Field label="Doctor Notes / डॉक्टर नोट्स">
                      <textarea style={{ ...inp, height: 80, resize: 'vertical' }}
                        placeholder="Additional observations…"
                        value={visitForm.doctor_notes}
                        onChange={(e) => setVisitForm({ ...visitForm, doctor_notes: e.target.value })} />
                    </Field>

                    <button type="submit" style={{ ...styles.btnPrimary, marginTop: 10 }}>
                      + Save Visit
                    </button>
                  </form>
                </Card>

                {/* Visit History */}
                <Card title={`Visit History (${visits.length})`} icon="🗂" accent="#7b5ea7">
                  <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 12 }} className="export-row">
                    <button type="button" style={{ ...styles.btnSecondary, background: '#7b5ea7' }} onClick={exportVisitsToExcel}>
                      📥 Export to Excel
                    </button>
                  </div>
                  {visits.length === 0 && (
                    <div style={styles.emptyState}>No visits recorded yet.</div>
                  )}
                  {visits.map((visit, i) => (
                    <div key={visit.visit_id} style={{ ...styles.visitCard, borderLeftColor: i % 2 === 0 ? '#2e6b5e' : '#c47c2b' }}>
                      <div style={styles.visitDate}>
                        📅 {new Date(visit.visit_date).toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' })}
                      </div>
                      <div style={styles.visitGrid} className="visit-grid">
                        {visit.symptoms?.length > 0 && (
                          <div>
                            <div style={styles.visitLbl}>लक्षणे (Symptoms)</div>
                            <div style={styles.tagRow}>
                              {visit.symptoms.map((s) => (
                                <span key={s} style={{ ...styles.tag, background: '#fff3e0', color: '#c47c2b', border: '1px solid #f0c980' }}>
                                  {s.replace('::', ' → ')}
                                </span>
                              ))}
                            </div>
                          </div>
                        )}
                        {visit.fee !== null && visit.fee !== '' && visit.fee !== undefined && (
                          <div>
                            <div style={styles.visitLbl}>फी (Fee Taken)</div>
                            <div style={styles.visitVal}>₹{visit.fee}</div>
                          </div>
                        )}
                        {visit.diagnosis && (
                          <div>
                            <div style={styles.visitLbl}>निदान (Diagnosis)</div>
                            <div style={styles.visitVal}>{visit.diagnosis}</div>
                          </div>
                        )}
                        {visit.medicines && (
                          <div>
                            <div style={styles.visitLbl}>औषधे (Medicines)</div>
                            <div style={styles.visitVal}>{visit.medicines}</div>
                          </div>
                        )}
                        {visit.doctor_notes && (
                          <div>
                            <div style={styles.visitLbl}>डॉक्टर नोट्स</div>
                            <div style={styles.visitVal}>{visit.doctor_notes}</div>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </Card>
              </>
            )}
          </div>
        )}
      </div>
    </>
  )
}

// ─── Style tokens ─────────────────────────────────────────────────────────────

const styles = {
  header: {
    background: 'linear-gradient(135deg, #1a4a40 0%, #2e6b5e 60%, #3d8b7a 100%)',
    padding: '0',
    boxShadow: '0 4px 20px #1a4a4044',
  },
  headerInner: {
    maxWidth: 1200,
    margin: 'auto',
    padding: '18px 32px',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  logo: { display: 'flex', alignItems: 'center', gap: 16 },
  logoIcon: { fontSize: 40, filter: 'drop-shadow(0 2px 4px #0005)' },
  logoTitle: { fontFamily: "'Playfair Display', serif", fontSize: 22, color: '#fff', fontWeight: 700, letterSpacing: '0.02em' },
  logoSub: { fontSize: 12, color: '#a8d5ca', marginTop: 2, letterSpacing: '0.05em' },
  headerStats: { display: 'flex', gap: 24 },
  stat: { display: 'flex', flexDirection: 'column', alignItems: 'center', background: '#ffffff18', padding: '8px 20px', borderRadius: 12 },
  statNum: { fontSize: 24, fontWeight: 700, color: '#fff', fontFamily: "'Playfair Display', serif" },
  statLbl: { fontSize: 11, color: '#a8d5ca', textTransform: 'uppercase', letterSpacing: '0.08em' },

  main: { maxWidth: 1200, margin: '0 auto', padding: '0 24px 60px' },

  tabs: { display: 'flex', gap: 0, borderBottom: '2px solid #e0d5c5', margin: '28px 0 0', },
  tab: { padding: '12px 28px', border: 'none', background: 'none', cursor: 'pointer', fontSize: 14, fontWeight: 500, color: '#9a8a7a', borderBottom: '2px solid transparent', marginBottom: -2, display: 'flex', alignItems: 'center', gap: 8, fontFamily: "'DM Sans', sans-serif" },
  tabActive: { color: '#2e6b5e', borderBottomColor: '#2e6b5e' },
  tabBadge: { background: '#2e6b5e22', color: '#2e6b5e', fontSize: 11, padding: '2px 8px', borderRadius: 20, fontWeight: 500 },

  grid: { display: 'flex', flexDirection: 'column', gap: 24, marginTop: 24 },

  card: { background: '#fff', borderRadius: 16, boxShadow: '0 2px 16px #3a2e2210', overflow: 'hidden' },
  cardHeader: { padding: '18px 24px 14px', display: 'flex', alignItems: 'center', borderBottom: '1px solid #f0e8d8' },
  cardTitle: { fontFamily: "'Playfair Display', serif", fontSize: 18, fontWeight: 600 },
  cardBody: { padding: '20px 24px' },

  field: { marginBottom: 16 },
  fieldLabel: { display: 'block', fontSize: 12, fontWeight: 500, color: '#7a6a5a', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 6 },

  twoCol: { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16, marginBottom: 0 },

  checkSection: { background: '#faf7f2', border: '1.5px solid #e8ddc8', borderRadius: 12, padding: '16px 18px', marginBottom: 18 },
  checkTitle: { fontFamily: "'Playfair Display', serif", fontSize: 15, color: '#3a2e22', marginBottom: 12, fontWeight: 600 },

  nestedItem: { background: '#fff', border: '1px solid #e8ddc8', borderRadius: 10, padding: '10px 14px', marginBottom: 8 },
  parentRow: { display: 'flex', alignItems: 'center', justifyContent: 'space-between' },
  checkLabel: { display: 'flex', alignItems: 'center', gap: 10, cursor: 'pointer', flex: 1 },
  customCheck: { width: 18, height: 18, borderRadius: 4, border: '2px solid #c9b99a', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0, cursor: 'pointer', transition: 'all 0.15s' },
  parentLabel: { fontSize: 15, fontWeight: 500, color: '#3a2e22' },
  expandBtn: { background: 'none', border: 'none', cursor: 'pointer', fontSize: 11, padding: '4px 8px', borderRadius: 6, fontWeight: 600 },
  subList: { marginTop: 10, paddingLeft: 28, display: 'flex', flexDirection: 'column', gap: 6 },
  subLabel: { display: 'flex', alignItems: 'center', gap: 8, cursor: 'pointer', padding: '4px 0' },
  subTags: { display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 8, paddingLeft: 28 },

  table: { width: '100%', borderCollapse: 'collapse', fontSize: 14 },
  th: { background: '#f5f0e8', padding: '12px 14px', textAlign: 'left', fontSize: 12, fontWeight: 600, color: '#7a6a5a', textTransform: 'uppercase', letterSpacing: '0.05em', borderBottom: '2px solid #e0d5c5' },
  td: { padding: '12px 14px', borderBottom: '1px solid #f0e8d8', color: '#3a2e22', verticalAlign: 'middle' },
  idBadge: { background: '#e8f5f2', color: '#2e6b5e', padding: '3px 8px', borderRadius: 20, fontWeight: 600, fontSize: 12 },
  genderBadge: { padding: '3px 10px', borderRadius: 20, fontSize: 12, fontWeight: 500 },
  actionBtn: { padding: '6px 12px', borderRadius: 7, border: '1.5px solid #e0d5c5', background: '#faf7f2', cursor: 'pointer', fontSize: 12, fontWeight: 500, color: '#5a4a3a', whiteSpace: 'nowrap' },

  btnPrimary: { background: 'linear-gradient(135deg, #2e6b5e, #3d8b7a)', color: '#fff', border: 'none', padding: '12px 24px', borderRadius: 10, cursor: 'pointer', fontWeight: 600, fontSize: 14, fontFamily: "'DM Sans', sans-serif", boxShadow: '0 4px 12px #2e6b5e44' },
  btnSecondary: { background: '#7b5ea7', color: '#fff', border: 'none', padding: '10px 18px', borderRadius: 9, cursor: 'pointer', fontWeight: 500, fontSize: 14, fontFamily: "'DM Sans', sans-serif" },
  btnGhost: { background: 'none', color: '#9a8a7a', border: '1.5px solid #e0d5c5', padding: '12px 20px', borderRadius: 10, cursor: 'pointer', fontSize: 14, fontFamily: "'DM Sans', sans-serif" },

  patientBanner: { background: 'linear-gradient(135deg, #e8f5f2, #f0faf7)', border: '1.5px solid #b8ddd6', borderRadius: 14, padding: '16px 24px', display: 'flex', alignItems: 'center', gap: 16 },

  visitCard: { borderLeft: '4px solid', borderRadius: 10, background: '#faf7f2', padding: '16px 18px', marginBottom: 14 },
  visitDate: { fontSize: 12, color: '#7a6a5a', marginBottom: 10, fontWeight: 500 },
  visitGrid: { display: 'grid', gap: 12 },
  visitLbl: { fontSize: 11, fontWeight: 600, color: '#9a8a7a', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 4 },
  visitVal: { fontSize: 14, color: '#3a2e22' },
  tagRow: { display: 'flex', flexWrap: 'wrap', gap: 6 },
  tag: { fontSize: 12, padding: '3px 10px', borderRadius: 20, fontWeight: 500 },

  emptyState: { textAlign: 'center', padding: '40px 20px', color: '#9a8a7a' },
}