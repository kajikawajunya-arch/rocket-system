import { useState, useMemo, useEffect } from 'react';
import { 
  AlertCircle, 
  CheckCircle2, 
  Clock, 
  Package, 
  MessageCircle,
  Truck,
  Filter,
  Calendar,
  Loader2
} from 'lucide-react';

import { ComparisonView } from './ComparisonView';

// GAS WebアプリのURL（新しくデプロイされたもの）
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbwAMsW-5Jir6SkJ6VykJvxt2KKO6PP6S3mqNAv1JBl8K9v6leDWYNebcj26mcASmZOKqw/exec';

type KomaFilter = 'ALL' | 'A' | 'B' | 'AB';

function App() {
  const [currentView, setCurrentView] = useState<'dashboard' | 'comparison'>('dashboard');
  const [komaFilter, setKomaFilter] = useState<KomaFilter>('ALL');
  const [showOnlyUnreported, setShowOnlyUnreported] = useState(false);
  
  // 今日の日付を初期値にセット (YYYY-MM-DD形式)
  const today = new Date();
  const todayStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const [selectedDate, setSelectedDate] = useState(todayStr);

  const [isLoading, setIsLoading] = useState(false);
  const [fetchedData, setFetchedData] = useState<any[]>([]);

  // 日付が変更されたら、該当月のデータをAPIから取得する
  useEffect(() => {
    const fetchSpreadsheetData = async () => {
      setIsLoading(true);
      try {
        const dateObj = new Date(selectedDate);
        const targetMonthStr = `${dateObj.getFullYear()}年${dateObj.getMonth() + 1}月`;
        
        // GASはリダイレクトを返すので redirect: 'follow' を明示
        const response = await fetch(`${GAS_API_URL}?month=${encodeURIComponent(targetMonthStr)}`, {
          method: 'GET',
          redirect: 'follow'
        });
        const json = await response.json();
        
        if (json.data) {
          setFetchedData(json.data);
        } else {
          setFetchedData([]);
        }
      } catch (error) {
        console.error("データの取得に失敗しました", error);
        // エラー時はアラートで知らせる
        alert("データの取得に失敗しました。URLや権限を確認してください。");
        setFetchedData([]);
      } finally {
        setIsLoading(false);
      }
    };

    fetchSpreadsheetData();
  }, [selectedDate]); // selectedDateが変わるたびに再取得（簡易的なキャッシュなし）

  // APIから取得したデータを、選択された日付（日のみ、例："8月1日"）で絞り込み、さらに状態を判定
  const processedData = useMemo(() => {
    const dateObj = new Date(selectedDate);
    const targetDayStr = `${dateObj.getMonth() + 1}月${dateObj.getDate()}日`; // "8月1日"

    // 該当する日付の行だけを抽出
    const daysData = fetchedData.filter(row => {
      const rowDate = row['日付'];
      if (!rowDate) return false;
      
      const strRowDate = String(rowDate);
      
      // 1. "8月1日" として完全・部分一致
      if (strRowDate.includes(targetDayStr)) return true;
      
      // 2. ISO文字列（"2026-07-31T15:00:00.000Z" 等）への対策
      // JSTの 8月1日は、UTCでは 7月31日 15:00 になるため、スプレッドシートから前日の日付として送られてくる
      const selD = new Date(selectedDate);
      
      // 当日の日付文字列 "2026-08-01"
      const todayStr = `${selD.getFullYear()}-${String(selD.getMonth() + 1).padStart(2, '0')}-${String(selD.getDate()).padStart(2, '0')}`;
      
      // 前日の日付文字列 "2026-07-31" (UTC表記対策)
      const prevD = new Date(selD.getTime() - 24 * 60 * 60 * 1000);
      const prevStr = `${prevD.getFullYear()}-${String(prevD.getMonth() + 1).padStart(2, '0')}-${String(prevD.getDate()).padStart(2, '0')}`;
      
      // rowDate が今日または前日のYYYY-MM-DDを含んでいれば一致とする
      if (strRowDate.includes(todayStr) || strRowDate.includes(prevStr)) {
        return true;
      }

      return false;
    });

    return daysData.map((row, index) => {
      // ※GAS側で未報告の場合、空文字になっていることを想定。右端の未確認列にエラーがある場合も拾う
      const errorDetail = row['エラー詳細'] || row['未確認ドライバー'] || '';
      // 報告済みかどうかの判定：配達件数が入っており、かつエラー詳細が空であること
      const hasReported = (row['配達件数'] !== '' && row['配達件数'] !== undefined && row['配達件数'] !== null) && !errorDetail;
      
      return {
        id: index.toString(),
        date: targetDayStr, // 表示用に整形された日付を使用
        koma: row['コマ数'] || '不明',
        codeName: row['コード名'] || '名称不明',
        isReported: hasReported,
        errorDetail: errorDetail,
        reportDetails: {
          workerName: row['稼働者(申告)'] || row['稼働者'] || row['コード名'] || '不明',
          deliveries: row['配達件数']
        }
      };
    });
  }, [fetchedData, selectedDate]);

  // フィルター適用後のデータ
  const filteredData = useMemo(() => {
    let result = processedData;
    
    // コマ数フィルター
    if (komaFilter !== 'ALL') {
      result = result.filter(d => d.koma === komaFilter);
    }
    
    // 未報告フィルター
    if (showOnlyUnreported) {
      result = result.filter(d => !d.isReported);
    }
    
    return result;
  }, [processedData, komaFilter, showOnlyUnreported]);

  // サマリー計算
  const totalCount = filteredData.length;
  const reportedCount = filteredData.filter(d => d.isReported).length;
  const unreportedCount = totalCount - reportedCount;

  // 照合画面が選択されている場合は ComparisonView をレンダリング
  if (currentView === 'comparison') {
    return (
      <ComparisonView onBack={() => setCurrentView('dashboard')} />
    );
  }

  return (
    <div className="app-container">
      {/* Header */}
      <header className="header">
        <h1>ロケットナウ報告管理</h1>
        <div className="header-actions" style={{ display: 'flex', gap: '1rem', alignItems: 'center' }}>
          <div style={{ position: 'relative', display: 'flex', alignItems: 'center', background: 'rgba(0,0,0,0.3)', padding: '0.4rem 0.8rem', borderRadius: '8px', border: '1px solid var(--card-border)' }}>
            <Calendar size={16} style={{ position: 'absolute', left: '0.8rem', pointerEvents: 'none', color: 'var(--text-secondary)' }} />
            <input 
              type="date" 
              value={selectedDate}
              onChange={(e) => setSelectedDate(e.target.value)}
              onClick={(e) => {
                // クリック時にカレンダーピッカーを強制的に開く
                try {
                  (e.target as HTMLInputElement).showPicker();
                } catch(err) {}
              }}
              style={{ background: 'transparent', border: 'none', color: 'var(--text-primary)', outline: 'none', fontFamily: 'inherit', fontSize: '0.9rem', paddingLeft: '1.8rem', cursor: 'pointer', width: '130px' }}
            />
          </div>
          <button className="action-btn" onClick={() => setCurrentView('comparison')}>
            <Filter size={18} />
            予実突合
          </button>
        </div>
      </header>

      {/* Main Content */}
      <main className="main-content">
        <div className="dashboard-grid">
          
          {/* Summary Section */}
          <div className="summary-container">
            <div className="summary-box">
              <span className="summary-label">稼働台数</span>
              <span className="summary-value">{totalCount}</span>
            </div>
            <div className="summary-box">
              <span className="summary-label">報告完了</span>
              <span className="summary-value" style={{ color: 'var(--status-reported-text)' }}>
                {reportedCount}
              </span>
            </div>
            <div className={`summary-box ${unreportedCount > 0 ? 'alert' : ''}`}>
              <span className="summary-label">未報告アラート</span>
              <span className="summary-value">{unreportedCount}</span>
            </div>
          </div>

          <div className="glass-card">
            {/* Filter Controls */}
            <div className="controls-bar">
              <div className="filter-group">
                <button 
                  className={`filter-btn ${komaFilter === 'ALL' ? 'active' : ''}`}
                  onClick={() => setKomaFilter('ALL')}
                >
                  すべて
                </button>
                <button 
                  className={`filter-btn ${komaFilter === 'A' ? 'active' : ''}`}
                  onClick={() => setKomaFilter('A')}
                >
                  Aコマ
                </button>
                <button 
                  className={`filter-btn ${komaFilter === 'B' ? 'active' : ''}`}
                  onClick={() => setKomaFilter('B')}
                >
                  Bコマ
                </button>
                <button 
                  className={`filter-btn ${komaFilter === 'AB' ? 'active' : ''}`}
                  onClick={() => setKomaFilter('AB')}
                >
                  AB通し
                </button>
              </div>
              
              <div style={{ marginLeft: '1rem', display: 'flex', alignItems: 'center' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '0.4rem', cursor: 'pointer', fontSize: '0.85rem', color: 'var(--text-primary)', background: 'rgba(255,255,255,0.05)', padding: '0.4rem 0.8rem', borderRadius: '8px', border: '1px solid var(--card-border)' }}>
                  <input 
                    type="checkbox" 
                    checked={showOnlyUnreported}
                    onChange={(e) => setShowOnlyUnreported(e.target.checked)}
                    style={{ accentColor: 'var(--status-unreported)' }}
                  />
                  未報告のみ表示
                </label>
              </div>

              <div style={{ marginLeft: 'auto', fontSize: '0.875rem', color: 'var(--text-secondary)' }}>
                {filteredData.length} 件を表示中
              </div>
            </div>

            {/* List */}
            <div className="driver-list">
              {isLoading ? (
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>
                  <Loader2 className="spinner" size={32} style={{ animation: 'spin 1s linear infinite', marginBottom: '1rem' }} />
                  <p>スプレッドシートからデータを読み込み中...</p>
                  <style>{`@keyframes spin { 100% { transform: rotate(360deg); } }`}</style>
                </div>
              ) : (
                <>
                  {filteredData.map((driver) => (
                    <div key={driver.id} className={`driver-item ${!driver.isReported ? 'unreported' : ''}`}>
                      <div className="driver-info">
                        <div className="driver-name-group">
                          <span className="driver-name">{driver.codeName}</span>
                          <span className="driver-code-badge">コード名</span>
                        </div>
                        
                        <div className="driver-meta">
                          <span style={{ display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--text-secondary)' }}>
                            <Calendar size={14} /> {driver.date}
                          </span>
                          <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                            <Clock size={14} /> {driver.koma}コマ
                          </span>
                          {driver.reportDetails && (
                            <>
                              <span style={{ display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--text-primary)' }}>
                                <Truck size={14} /> 稼働者: {driver.reportDetails.workerName}
                              </span>
                              <span style={{ display: 'flex', alignItems: 'center', gap: '4px', color: driver.isReported ? 'var(--status-reported-text)' : 'var(--text-secondary)' }}>
                                <Package size={14} /> {driver.reportDetails.deliveries ? `${driver.reportDetails.deliveries}件` : '未報告'}
                              </span>
                            </>
                          )}
                          {!driver.isReported && (
                            <span style={{ display: 'flex', alignItems: 'flex-start', gap: '4px', color: 'var(--status-unreported-text)', fontSize: '0.8rem', marginTop: '4px' }}>
                              <AlertCircle size={12} style={{ marginTop: '2px', flexShrink: 0 }} /> 
                              <span style={{ wordBreak: 'break-all' }}>
                                {driver.errorDetail ? `※${driver.errorDetail.split('(リスト:')[0]}` : '※配達件数がまだ報告されていません'}
                              </span>
                            </span>
                          )}
                        </div>
                      </div>

                      <div className="driver-status">
                        {driver.isReported ? (
                          <span className="status-badge reported">
                            <CheckCircle2 size={14} /> 報告済
                          </span>
                        ) : (
                          <>
                            <span className="status-badge unreported">
                              <AlertCircle size={14} /> 未報告
                            </span>
                            <button className="action-btn primary" title="LINEでリマインドを送る（モック）" onClick={() => alert('LINEへの自動リマインド送信機能は準備中です')}>
                              <MessageCircle size={16} />
                              リマインド
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  ))}
                  
                  {filteredData.length === 0 && (
                    <div style={{ textAlign: 'center', padding: '3rem', color: 'var(--text-secondary)' }}>
                      該当する稼働データがありません
                      <br /><br />
                      <small>選択中の日付: {selectedDate}</small>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;
