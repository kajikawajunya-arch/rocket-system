import React, { useState, useMemo } from 'react';
import { ArrowLeft, Upload, FileSpreadsheet, CheckCircle2, AlertTriangle, Info, Calendar, Link as LinkIcon, Trash2 } from 'lucide-react';
import * as XLSX from 'xlsx';

// --- Types ---
interface ParsedRecord {
  id: string;
  date: string; // YYYY-MM-DD
  rawName: string;
  normName: string;
  koma: string;
  count: number | null;
}

interface MappedGroup {
  id: string;
  payment: ParsedRecord;
  shifts: ParsedRecord[];
}

export const ComparisonView: React.FC<{ onBack: () => void }> = ({ onBack }) => {
  // File states
  const [shiftFileName, setShiftFileName] = useState<string | null>(null);
  const [paymentFileName, setPaymentFileName] = useState<string | null>(null);
  
  // Data states
  const [shiftRecords, setShiftRecords] = useState<ParsedRecord[]>([]);
  const [paymentRecords, setPaymentRecords] = useState<ParsedRecord[]>([]);
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  // Mapping state (手動で紐付けた代走グループ)
  const [mappedGroups, setMappedGroups] = useState<MappedGroup[]>([]);
  const [selectedPaymentId, setSelectedPaymentId] = useState<string | null>(null);
  const [selectedShiftIds, setSelectedShiftIds] = useState<Set<string>>(new Set());

  // --- Helpers ---
  const normalizeDate = (rawDate: any): string | null => {
    if (!rawDate) return null;
    if (typeof rawDate === 'number') {
      const date = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
      return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
    }
    if (typeof rawDate === 'string') {
      const match = rawDate.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
      if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
    }
    return null;
  };

  const normalizeName = (name: any): string => String(name || '').replace(/\s+/g, '');

  // 柔軟なExcelパーサー
  const parseExcelFile = async (
    file: File, 
    type: 'shift' | 'payment'
  ): Promise<ParsedRecord[]> => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array', cellDates: true });
    
    // シート選択：2番目(インデックス1)を優先、なければ1番目(インデックス0)
    const sheetIdx = workbook.SheetNames.length > 1 ? 1 : 0;
    const worksheet = workbook.Sheets[workbook.SheetNames[sheetIdx]];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    let headerRowIndex = -1;
    let colDate = -1, colName = -1, colKoma = -1, colCount = -1;

    // ヘッダー行の探索（表記揺れ対応）
    for (let i = 0; i < Math.min(rows.length, 50); i++) {
      const row = rows[i];
      if (!Array.isArray(row)) continue;
      const rowStrs = row.map(cell => String(cell || '').replace(/\s+/g, ''));
      
      const hasDate = rowStrs.some(s => s.includes('日付') || s.includes('勤務日') || s.includes('稼働日'));
      const hasName = rowStrs.some(s => s.includes('名前') || s.includes('ドライバー') || s.includes('フルネーム') || s.includes('氏名'));
      
      if (hasDate && hasName) {
        headerRowIndex = i;
        colDate = rowStrs.findIndex(s => s.includes('日付') || s.includes('勤務日') || s.includes('稼働日'));
        colName = rowStrs.findIndex(s => s.includes('名前') || s.includes('ドライバー') || s.includes('フルネーム') || s.includes('氏名'));
        colKoma = rowStrs.findIndex(s => s.includes('稼働') || s.includes('ブロック') || s.includes('コマ') || s.includes('シフト'));
        
        if (type === 'payment') {
          colCount = rowStrs.findIndex(s => s === '件数' || s.includes('完了件数') || s.includes('配達件数'));
        }
        break;
      }
    }

    if (headerRowIndex === -1 || colDate === -1 || colName === -1 || (type === 'payment' && colCount === -1)) {
      throw new Error(`${file.name}: 必要な列が見つかりません。フォーマットを確認してください。`);
    }

    const records: ParsedRecord[] = [];
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;
      
      const rawName = row[colName];
      if (!rawName || typeof rawName !== 'string' || rawName.includes('小計') || rawName.includes('合計')) continue;
      
      const rawDate = row[colDate];
      const dateStr = normalizeDate(rawDate);
      if (!dateStr) continue;

      const normName = normalizeName(rawName);
      const koma = colKoma !== -1 ? String(row[colKoma] || '') : '';
      const count = type === 'payment' ? (parseFloat(row[colCount]) || 0) : null;

      records.push({
        id: `${type}-${dateStr}-${normName}-${i}`,
        date: dateStr,
        rawName,
        normName,
        koma,
        count
      });
    }
    return records;
  };

  // --- Handlers ---
  const handleShiftUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setIsProcessing(true);
      const records = await parseExcelFile(file, 'shift');
      setShiftRecords(records);
      setShiftFileName(file.name);
    } catch (err: any) {
      alert(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handlePaymentUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      setIsProcessing(true);
      const records = await parseExcelFile(file, 'payment');
      setPaymentRecords(records);
      setPaymentFileName(file.name);
    } catch (err: any) {
      alert(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleMapSelected = () => {
    if (!selectedPaymentId || selectedShiftIds.size === 0) return;
    
    const payment = paymentRecords.find(p => p.id === selectedPaymentId);
    const shifts = shiftRecords.filter(s => selectedShiftIds.has(s.id));
    
    if (payment && shifts.length > 0) {
      setMappedGroups(prev => [...prev, {
        id: `group-${Date.now()}`,
        payment,
        shifts
      }]);
      // リセット
      setSelectedPaymentId(null);
      setSelectedShiftIds(new Set());
    }
  };

  const handleUnmap = (groupId: string) => {
    setMappedGroups(prev => prev.filter(g => g.id !== groupId));
  };

  const toggleShiftSelection = (id: string) => {
    const newSet = new Set(selectedShiftIds);
    if (newSet.has(id)) newSet.delete(id);
    else newSet.add(id);
    setSelectedShiftIds(newSet);
  };

  // --- Helper: 同日・同ドライバーの複数行を1つに集約（コマ表記合算 A1+B1等） ---
  const groupRecordsByPerson = (records: ParsedRecord[]): ParsedRecord[] => {
    const map = new Map<string, ParsedRecord[]>();
    records.forEach(r => {
      const list = map.get(r.normName) || [];
      list.push(r);
      map.set(r.normName, list);
    });

    const grouped: ParsedRecord[] = [];
    map.forEach((list, normName) => {
      const first = list[0];
      
      // コマの集約：重複除去してソート後、"+" で結合（例: A1 と B1 -> A1+B1）
      const komaSet = new Set<string>();
      list.forEach(item => {
        if (item.koma) {
          item.koma.split(/[\+,\s]+/).forEach(k => {
            if (k.trim()) komaSet.add(k.trim());
          });
        }
      });
      
      const sortedKomas = Array.from(komaSet).sort();
      const combinedKoma = sortedKomas.length > 0 ? sortedKomas.join('+') : first.koma;

      // 件数の集約
      let totalCount: number | null = null;
      let hasCount = false;
      list.forEach(item => {
        if (item.count !== null && item.count !== undefined) {
          totalCount = (totalCount || 0) + item.count;
          hasCount = true;
        }
      });

      grouped.push({
        id: list.map(l => l.id).join('__'),
        date: first.date,
        rawName: first.rawName,
        normName: normName,
        koma: combinedKoma,
        count: hasCount ? totalCount : null
      });
    });

    return grouped;
  };

  // --- Analysis Logic ---
  const analysisResult = useMemo(() => {
    if (!selectedDate) return null;

    // 選択された日付のレコードを取得し、ドライバー単位で同日データ・複数コマを1つに統合（A1+B1など）
    const rawShifts = shiftRecords.filter(s => s.date === selectedDate);
    const rawPayments = paymentRecords.filter(p => p.date === selectedDate);
    
    const shifts = groupRecordsByPerson(rawShifts);
    const payments = groupRecordsByPerson(rawPayments);
    
    // マッピング済みのIDを収集
    const mappedPaymentIds = new Set<string>();
    const mappedShiftIds = new Set<string>();
    mappedGroups.forEach(g => {
      mappedPaymentIds.add(g.payment.id);
      g.shifts.forEach(s => mappedShiftIds.add(s.id));
    });

    const exactMatch: { payment: ParsedRecord, shift: ParsedRecord }[] = [];
    const unmatchedPayments: ParsedRecord[] = []; // 上位店のみ
    const unmatchedShifts: ParsedRecord[] = []; // 自社のみ

    // 支払通知書ベースで突き合わせ
    const matchedShiftIds = new Set<string>();
    
    payments.forEach(p => {
      if (mappedPaymentIds.has(p.id)) return; // マッピング済みはスキップ

      // シフト側に同じ名前がいるか？
      const s = shifts.find(sh => sh.normName === p.normName && !mappedShiftIds.has(sh.id) && !matchedShiftIds.has(sh.id));
      if (s) {
        exactMatch.push({ payment: p, shift: s });
        matchedShiftIds.add(s.id);
      } else {
        unmatchedPayments.push(p);
      }
    });

    // 余ったシフト
    shifts.forEach(s => {
      if (!mappedShiftIds.has(s.id) && !matchedShiftIds.has(s.id)) {
        unmatchedShifts.push(s);
      }
    });

    // 該当日のマッピンググループ
    const todaysGroups = mappedGroups.filter(g => g.payment.date === selectedDate);

    return {
      exactMatch,
      unmatchedPayments,
      unmatchedShifts,
      todaysGroups
    };
  }, [shiftRecords, paymentRecords, selectedDate, mappedGroups]);

  // 利用可能な日付リスト
  const availableDates = useMemo(() => {
    const dates = new Set([...shiftRecords.map(r=>r.date), ...paymentRecords.map(r=>r.date)]);
    const arr = Array.from(dates).sort();
    if (arr.length > 0 && !selectedDate) setSelectedDate(arr[0]);
    return arr;
  }, [shiftRecords, paymentRecords]);

  // --- UI ---
  return (
    <div className="comparison-container" style={{ padding: '2rem', maxWidth: '1400px', margin: '0 auto', color: 'var(--text-primary)' }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'center', marginBottom: '2rem' }}>
        <button onClick={onBack} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', background: 'transparent', border: 'none', color: 'var(--accent-primary)', cursor: 'pointer', fontSize: '1rem', padding: 0 }}>
          <ArrowLeft size={20} />
          ダッシュボードへ戻る
        </button>
        <h1 style={{ marginLeft: '1.5rem', fontSize: '1.5rem', margin: 0 }}>シフト予実突合（2ファイル比較・代走マッピング）</h1>
      </div>

      {/* Uploaders */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '2rem', marginBottom: '2rem' }}>
        {/* 自社シフト Upload */}
        <div className="glass-card" style={{ padding: '2rem', textAlign: 'center', borderStyle: 'dashed', borderWidth: '2px', borderColor: shiftFileName ? 'var(--status-reported)' : 'var(--card-border)' }}>
          <input 
            type="file" 
            accept=".xlsx, .xls" 
            id="shift-upload" 
            style={{ display: 'none' }} 
            onClick={(e) => { (e.currentTarget as HTMLInputElement).value = ''; }}
            onChange={handleShiftUpload} 
          />
          <label htmlFor="shift-upload" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem' }}>
            <div style={{ background: shiftFileName ? 'rgba(46, 204, 113, 0.1)' : 'rgba(255,255,255,0.05)', padding: '1rem', borderRadius: '50%' }}>
              <Upload size={32} color={shiftFileName ? 'var(--status-reported)' : 'var(--accent-primary)'} />
            </div>
            <div>
              <h3 style={{ margin: '0 0 0.5rem 0' }}>1. 自社提出シフト予定</h3>
              {shiftFileName ? (
                <div style={{ color: 'var(--status-reported-text)', fontWeight: 'bold' }}>✅ {shiftFileName}</div>
              ) : (
                <p style={{ margin: 0, color: 'var(--text-secondary)', fontSize: '0.9rem' }}>shift_template_*.xlsx を選択</p>
              )}
            </div>
          </label>
        </div>

        {/* 支払通知書 Upload */}
        <div className="glass-card" style={{ padding: '2rem', textAlign: 'center', borderStyle: 'dashed', borderWidth: '2px', borderColor: paymentFileName ? 'var(--status-reported)' : 'var(--card-border)' }}>
          <input 
            type="file" 
            accept=".xlsx, .xls" 
            id="payment-upload" 
            style={{ display: 'none' }} 
            onClick={(e) => { (e.currentTarget as HTMLInputElement).value = ''; }}
            onChange={handlePaymentUpload} 
          />
          <label htmlFor="payment-upload" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '1rem' }}>
            <div style={{ background: paymentFileName ? 'rgba(46, 204, 113, 0.1)' : 'rgba(255,255,255,0.05)', padding: '1rem', borderRadius: '50%' }}>
              <Upload size={32} color={paymentFileName ? 'var(--status-reported)' : 'var(--accent-primary)'} />
            </div>
            <div>
              <h3 style={{ margin: '0 0 0.5rem 0' }}>2. 上位店 支払通知書</h3>
              {paymentFileName ? (
                <div style={{ color: 'var(--status-reported-text)', fontWeight: 'bold' }}>✅ {paymentFileName}</div>
              ) : (
                <p style={{ margin: 0, color: 'var(--text-secondary)', fontSize: '0.9rem' }}>代理店支払通知書_*.xlsx を選択</p>
              )}
            </div>
          </label>
        </div>
      </div>

      {isProcessing && <div style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>Excelを解析中...</div>}

      {/* Results Area */}
      {shiftFileName && paymentFileName && analysisResult && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }}>
          
          {/* Controls */}
          <div className="glass-card" style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '1rem 1.5rem' }}>
            <h2 style={{ margin: 0, fontSize: '1.2rem' }}>照合結果・代走マッピング</h2>
            <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
              <Calendar size={18} color="var(--text-secondary)" />
              <select 
                value={selectedDate || ''} 
                onChange={e => setSelectedDate(e.target.value)}
                style={{ background: 'var(--card-bg)', border: '1px solid var(--card-border)', color: 'var(--text-primary)', padding: '0.5rem', borderRadius: '8px', outline: 'none' }}
              >
                {availableDates.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>
          </div>

          {/* Puzzle UI (Left: Payment, Right: Shift) */}
          <div style={{ display: 'flex', gap: '1.5rem', flexWrap: 'wrap' }}>
            
            {/* Left: 支払通知書のみ (未確定の実績) */}
            <div className="glass-card" style={{ flex: '1 1 350px', display: 'flex', flexDirection: 'column' }}>
              <div style={{ padding: '1rem', borderBottom: '1px solid var(--card-border)', background: 'rgba(231, 76, 60, 0.05)' }}>
                <h3 style={{ margin: 0, color: 'var(--status-unreported-text)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <AlertTriangle size={18} /> 上位店実績のみに存在 (代走先？)
                </h3>
              </div>
              <div style={{ padding: '1rem', flex: 1, overflowY: 'auto', maxHeight: '400px' }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>シフト表にいないのに実績があります。右のリストから代走した人を選んで紐付けてください。</p>
                {analysisResult.unmatchedPayments.length === 0 ? (
                  <div style={{ textAlign: 'center', color: 'var(--text-secondary)', padding: '2rem 0' }}>該当なし</div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                    {analysisResult.unmatchedPayments.map(p => (
                      <div 
                        key={p.id} 
                        onClick={() => setSelectedPaymentId(selectedPaymentId === p.id ? null : p.id)}
                        style={{ 
                          padding: '0.75rem', borderRadius: '8px', cursor: 'pointer', transition: 'all 0.2s',
                          background: selectedPaymentId === p.id ? 'rgba(231, 76, 60, 0.2)' : 'rgba(255,255,255,0.05)',
                          border: `1px solid ${selectedPaymentId === p.id ? 'var(--status-unreported)' : 'transparent'}`
                        }}
                      >
                        <div style={{ fontWeight: 'bold' }}>{p.rawName}</div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          <span>シフト: {p.koma}</span>
                          <span style={{ color: 'var(--status-unreported-text)', fontWeight: 'bold' }}>{p.count} 件</span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {/* Link Action Button */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', flexDirection: 'column', gap: '1rem' }}>
              <button 
                onClick={handleMapSelected}
                disabled={!selectedPaymentId || selectedShiftIds.size === 0}
                className="action-btn primary"
                style={{ 
                  opacity: (!selectedPaymentId || selectedShiftIds.size === 0) ? 0.5 : 1,
                  pointerEvents: (!selectedPaymentId || selectedShiftIds.size === 0) ? 'none' : 'auto',
                  padding: '1rem', borderRadius: '50%', height: '60px', width: '60px', display: 'flex', justifyContent: 'center'
                }}
                title="選択した人を紐付ける（代走確定）"
              >
                <LinkIcon size={24} />
              </button>
              <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', textAlign: 'center', width: '80px' }}>
                選んで<br/>紐付け
              </div>
            </div>

            {/* Right: 自社シフトのみ (未確定の予定) */}
            <div className="glass-card" style={{ flex: '1 1 350px', display: 'flex', flexDirection: 'column' }}>
              <div style={{ padding: '1rem', borderBottom: '1px solid var(--card-border)', background: 'rgba(52, 152, 219, 0.05)' }}>
                <h3 style={{ margin: 0, color: 'var(--accent-primary)', display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <Info size={18} /> 自社シフトのみに存在 (欠勤？代走元？)
                </h3>
              </div>
              <div style={{ padding: '1rem', flex: 1, overflowY: 'auto', maxHeight: '400px' }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginBottom: '1rem' }}>提出したシフトにいるのに実績がありません。左の人の代わりに走った場合は選んでください（複数可）。</p>
                {analysisResult.unmatchedShifts.length === 0 ? (
                  <div style={{ textAlign: 'center', color: 'var(--text-secondary)', padding: '2rem 0' }}>該当なし</div>
                ) : (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem' }}>
                    {analysisResult.unmatchedShifts.map(s => (
                      <div 
                        key={s.id} 
                        onClick={() => toggleShiftSelection(s.id)}
                        style={{ 
                          padding: '0.75rem', borderRadius: '8px', cursor: 'pointer', transition: 'all 0.2s',
                          background: selectedShiftIds.has(s.id) ? 'rgba(52, 152, 219, 0.2)' : 'rgba(255,255,255,0.05)',
                          border: `1px solid ${selectedShiftIds.has(s.id) ? 'var(--accent-primary)' : 'transparent'}`
                        }}
                      >
                        <div style={{ fontWeight: 'bold' }}>{s.rawName}</div>
                        <div style={{ fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                          コマ: {s.koma}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>

          </div>

          {/* Mapped Groups (確定済みの代走) */}
          {analysisResult.todaysGroups.length > 0 && (
            <div className="glass-card">
              <div style={{ padding: '1rem', borderBottom: '1px solid var(--card-border)' }}>
                <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--status-reported-text)' }}>
                  <LinkIcon size={18} /> 手動で紐付け（代走確定）したリスト
                </h3>
              </div>
              <div style={{ padding: '1rem', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
                {analysisResult.todaysGroups.map(g => (
                  <div key={g.id} style={{ background: 'rgba(255,255,255,0.02)', border: '1px dashed var(--card-border)', borderRadius: '8px', padding: '1rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '2rem', flex: 1 }}>
                      {/* 上位店側 */}
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>上位店実績 (計 {g.payment.count}件)</div>
                        <div style={{ fontWeight: 'bold', fontSize: '1.1rem', color: 'var(--status-unreported-text)' }}>{g.payment.rawName} <span style={{ fontSize: '0.9rem', fontWeight: 'normal' }}>({g.payment.koma})</span></div>
                      </div>
                      
                      <ArrowLeft size={24} color="var(--text-secondary)" />
                      
                      {/* シフト側 */}
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: '0.8rem', color: 'var(--text-secondary)' }}>実際に走った人（自社シフト）</div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', marginTop: '4px' }}>
                          {g.shifts.map(s => (
                            <div key={s.id} style={{ fontWeight: 'bold', color: 'var(--accent-primary)' }}>{s.rawName} <span style={{ fontSize: '0.85rem', fontWeight: 'normal' }}>({s.koma})</span></div>
                          ))}
                        </div>
                      </div>
                    </div>
                    <button 
                      onClick={() => handleUnmap(g.id)}
                      style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', padding: '0.5rem' }}
                      title="紐付けを解除"
                    >
                      <Trash2 size={18} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Exact Matches (予定通り) */}
          <div className="glass-card">
            <div style={{ padding: '1rem', borderBottom: '1px solid var(--card-border)' }}>
              <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <CheckCircle2 size={18} color="var(--status-reported-text)" /> 完全一致（シフト予定通りに実績あり）: {analysisResult.exactMatch.length}名
              </h3>
            </div>
            <div style={{ padding: '1rem' }}>
              {analysisResult.exactMatch.length === 0 ? (
                <div style={{ textAlign: 'center', color: 'var(--text-secondary)' }}>該当なし</div>
              ) : (
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))', gap: '1rem' }}>
                  {analysisResult.exactMatch.map((match, idx) => (
                    <div key={idx} style={{ background: 'rgba(255,255,255,0.02)', padding: '0.75rem', borderRadius: '6px', border: '1px solid rgba(46, 204, 113, 0.2)' }}>
                      <div style={{ fontWeight: 'bold' }}>{match.shift.rawName}</div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.85rem', color: 'var(--text-secondary)', marginTop: '4px' }}>
                        <span>シフト: {match.shift.koma}</span>
                        <span style={{ color: 'var(--status-reported-text)' }}>実績: {match.payment.count}件</span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

        </div>
      )}
    </div>
  );
};
