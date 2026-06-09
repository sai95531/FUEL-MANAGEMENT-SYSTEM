<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Vignesh Fuel Station v3.1 — Complete Management System</title>
<link href="https://fonts.googleapis.com/css2?family=Rajdhani:wght@500;600;700&family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js"></script>
<style>
:root{
  --bg:#060a10;--surface:#0c1420;--surface2:#121c2e;--surface3:#18243a;
  --border:rgba(255,255,255,.07);--border2:rgba(255,255,255,.14);
  --text:#e2e8f0;--text-muted:#475569;--text-dim:#94a3b8;
  --petrol:#22d3ee;--petrol-dim:rgba(34,211,238,.12);--petrol-glow:rgba(34,211,238,.3);
  --diesel:#f97316;--diesel-dim:rgba(249,115,22,.12);--diesel-glow:rgba(249,115,22,.3);
  --green:#10b981;--green-dim:rgba(16,185,129,.12);
  --red:#ef4444;--red-dim:rgba(239,68,68,.12);
  --amber:#f59e0b;--amber-dim:rgba(245,158,11,.12);
  --blue:#3b82f6;--blue-dim:rgba(59,130,246,.12);
  --purple:#8b5cf6;--purple-dim:rgba(139,92,246,.12);
  --gold:#ffd700;
}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden;}
::-webkit-scrollbar{width:4px;height:4px;}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,.1);border-radius:2px;}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
@keyframes glow{0%,100%{box-shadow:0 0 8px var(--petrol-glow);}50%{box-shadow:0 0 24px var(--petrol-glow),0 0 48px rgba(34,211,238,.1);}}
@keyframes spin{to{transform:rotate(360deg);}}
@keyframes float{0%,100%{transform:translateY(0);}50%{transform:translateY(-6px);}}
@keyframes countUp{from{transform:translateY(8px);opacity:0;}to{transform:translateY(0);opacity:1;}}
@keyframes shimmer{0%{background-position:-200% 0;}100%{background-position:200% 0;}}

.header{background:linear-gradient(135deg,rgba(12,20,32,.98) 0%,rgba(18,28,46,.95) 100%);border-bottom:1px solid var(--border);padding:0 28px;position:sticky;top:0;z-index:100;backdrop-filter:blur(20px);}
.header-inner{display:flex;align-items:center;justify-content:space-between;height:64px;gap:16px;}
.logo-wrap{display:flex;align-items:center;gap:12px;}
.logo-icon{width:44px;height:44px;background:linear-gradient(135deg,var(--petrol),var(--diesel));border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:22px;flex-shrink:0;animation:float 4s infinite ease-in-out;}
.logo-text{font-family:'Rajdhani',sans-serif;}
.logo-name{font-size:18px;font-weight:700;color:var(--petrol);letter-spacing:2px;line-height:1.1;}
.logo-sub{font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--text-muted);}
.header-middle{display:flex;align-items:center;gap:24px;flex:1;}
.outlet-selector{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:6px 12px;cursor:pointer;font-size:12px;color:var(--text);min-width:180px;}
.header-clock{font-family:'DM Mono',monospace;font-size:13px;color:var(--amber);background:var(--amber-dim);padding:6px 14px;border-radius:8px;border:1px solid rgba(245,158,11,.2);text-align:center;}
.clock-date{font-size:9px;color:var(--text-muted);display:block;margin-bottom:2px;letter-spacing:1px;}
.live-dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:var(--green);animation:pulse 2s infinite;margin-right:6px;}
.stock-warning{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:8px;font-size:11px;font-weight:700;background:var(--red-dim);border:1px solid rgba(239,68,68,.3);color:var(--red);}

.user-switcher{display:flex;gap:6px;align-items:center;}
.user-btn{padding:6px 14px;border-radius:20px;border:1px solid var(--border);background:var(--surface2);color:var(--text-dim);cursor:pointer;font-size:11px;font-weight:600;transition:all .2s;display:flex;align-items:center;gap:6px;}
.user-btn:hover{border-color:var(--border2);color:var(--text);}
.user-btn.active{background:linear-gradient(135deg,var(--petrol-dim),rgba(34,211,238,.06));border-color:rgba(34,211,238,.3);color:var(--petrol);}
.user-btn.admin.active{background:linear-gradient(135deg,var(--amber-dim),rgba(245,158,11,.06));border-color:rgba(245,158,11,.3);color:var(--amber);}
.user-btn.cashier.active{background:linear-gradient(135deg,var(--purple-dim),rgba(139,92,246,.06));border-color:rgba(139,92,246,.3);color:var(--purple);}

.nav-wrap{background:var(--surface);border-bottom:1px solid var(--border);padding:0 28px;overflow-x:auto;}
.nav-tabs{display:flex;gap:0;white-space:nowrap;}
.nav-tab{padding:14px 20px;border-bottom:2px solid transparent;cursor:pointer;color:var(--text-muted);font-size:13px;font-weight:500;transition:all .2s;display:flex;align-items:center;gap:8px;user-select:none;position:relative;}
.nav-tab:hover{color:var(--text);background:rgba(255,255,255,.02);}
.nav-tab.active{color:var(--petrol);border-bottom-color:var(--petrol);}
.nav-tab .tab-icon{font-size:16px;}

.main{padding:24px 28px;max-width:1600px;margin:0 auto;}
.page{display:none;animation:fadeIn .3s ease;}
.page.active{display:block;}

.stats-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin-bottom:24px;}
.stat-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:18px 20px;position:relative;overflow:hidden;transition:all .2s;cursor:default;}
.stat-card:hover{transform:translateY(-3px);box-shadow:0 12px 32px rgba(0,0,0,.3);border-color:var(--border2);}
.stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;}
.stat-card.petrol::before{background:linear-gradient(90deg,var(--petrol),#67e8f9);}
.stat-card.diesel::before{background:linear-gradient(90deg,var(--diesel),#fbbf24);}
.stat-card.green::before{background:linear-gradient(90deg,var(--green),#34d399);}
.stat-card.amber::before{background:linear-gradient(90deg,var(--amber),#fb923c);}
.stat-card.blue::before{background:linear-gradient(90deg,var(--blue),#60a5fa);}
.stat-card.red::before{background:linear-gradient(90deg,var(--red),#f87171);}
.stat-card.purple::before{background:linear-gradient(90deg,var(--purple),#a78bfa);}
.stat-label{font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--text-muted);margin-bottom:8px;}
.stat-value{font-family:'Rajdhani',sans-serif;font-size:28px;font-weight:700;line-height:1;animation:countUp .4s ease;}
.stat-card.petrol .stat-value{color:var(--petrol);}
.stat-card.diesel .stat-value{color:var(--diesel);}
.stat-card.green .stat-value{color:var(--green);}
.stat-card.amber .stat-value{color:var(--amber);}
.stat-card.blue .stat-value{color:var(--blue);}
.stat-card.red .stat-value{color:var(--red);}
.stat-card.purple .stat-value{color:var(--purple);}
.stat-sub{font-size:11px;color:var(--text-muted);margin-top:4px;}
.stat-icon{position:absolute;right:16px;top:16px;font-size:24px;opacity:.2;}

.fuel-gauges{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:24px;}
.gauge-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:20px;position:relative;overflow:hidden;}
.gauge-title{font-family:'Rajdhani',sans-serif;font-size:13px;font-weight:600;letter-spacing:2px;text-transform:uppercase;margin-bottom:14px;display:flex;align-items:center;gap:8px;}
.gauge-title.petrol-t{color:var(--petrol);}
.gauge-title.diesel-t{color:var(--diesel);}
.gauge-bar-wrap{background:var(--surface2);border-radius:10px;height:28px;overflow:hidden;position:relative;border:1px solid var(--border);}
.gauge-bar{height:100%;border-radius:10px;transition:width 1s cubic-bezier(.4,0,.2,1);position:relative;overflow:hidden;}
.gauge-bar::after{content:'';position:absolute;inset:0;background:linear-gradient(90deg,transparent,rgba(255,255,255,.15),transparent);background-size:200% 100%;animation:shimmer 2s infinite;}
.gauge-bar.petrol-bar{background:linear-gradient(90deg,#0891b2,var(--petrol),#67e8f9);}
.gauge-bar.diesel-bar{background:linear-gradient(90deg,#c2410c,var(--diesel),#fbbf24);}
.gauge-info{display:flex;justify-content:space-between;margin-top:8px;font-size:12px;color:var(--text-dim);}
.gauge-liters{font-family:'Rajdhani',sans-serif;font-size:22px;font-weight:700;}
.gauge-liters.petrol-t{color:var(--petrol);}
.gauge-liters.diesel-t{color:var(--diesel);}
.gauge-pct{font-family:'DM Mono',monospace;font-size:20px;font-weight:600;color:var(--text-dim);}

.panel{background:var(--surface);border:1px solid var(--border);border-radius:14px;overflow:hidden;margin-bottom:20px;}
.panel-header{padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;}
.panel-title{font-family:'Rajdhani',sans-serif;font-size:12px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:var(--text-dim);}
.panel-body{padding:20px;}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px;}
.three-col{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;}

.dispense-wrap{max-width:600px;margin:0 auto;}
.fuel-type-select{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:22px;}
.fuel-type-btn{padding:22px;border-radius:12px;border:2px solid var(--border);background:var(--surface2);cursor:pointer;text-align:center;transition:all .2s;user-select:none;}
.fuel-type-btn:hover{transform:translateY(-2px);}
.fuel-type-btn .ftype-icon{font-size:36px;margin-bottom:8px;display:block;}
.fuel-type-btn .ftype-name{font-family:'Rajdhani',sans-serif;font-size:18px;font-weight:700;letter-spacing:1px;}
.fuel-type-btn .ftype-price{font-size:11px;color:var(--text-muted);margin-top:4px;}
.fuel-type-btn.sel-petrol{border-color:var(--petrol);background:var(--petrol-dim);box-shadow:0 4px 20px var(--petrol-glow);}
.fuel-type-btn.sel-petrol .ftype-name{color:var(--petrol);}
.fuel-type-btn.sel-diesel{border-color:var(--diesel);background:var(--diesel-dim);box-shadow:0 4px 20px var(--diesel-glow);}
.fuel-type-btn.sel-diesel .ftype-name{color:var(--diesel);}
.amount-input-wrap{position:relative;margin-bottom:18px;}
.amount-prefix{position:absolute;left:16px;top:50%;transform:translateY(-50%);font-family:'Rajdhani',sans-serif;font-size:24px;font-weight:700;color:var(--amber);}
.amount-input{width:100%;background:var(--surface2);border:2px solid var(--border);border-radius:12px;padding:18px 18px 18px 48px;color:var(--text);font-size:28px;font-family:'Rajdhani',sans-serif;font-weight:700;outline:none;transition:border-color .2s,box-shadow .2s;}
.amount-input:focus{border-color:var(--amber);box-shadow:0 0 0 4px rgba(245,158,11,.12);}
.liters-preview{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:16px 20px;margin-bottom:18px;display:flex;justify-content:space-between;align-items:center;}
.liters-label{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:var(--text-muted);}
.liters-val{font-family:'Rajdhani',sans-serif;font-size:32px;font-weight:700;color:var(--green);}
.dispense-btn{width:100%;background:linear-gradient(135deg,var(--green),#059669);color:#fff;border:none;border-radius:12px;padding:18px;font-size:16px;font-weight:700;font-family:'Rajdhani',sans-serif;letter-spacing:3px;text-transform:uppercase;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:10px;}
.dispense-btn:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(16,185,129,.4);}
.dispense-btn:active{transform:scale(.98);}
.quick-amounts{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;}
.quick-btn{padding:8px 16px;border-radius:8px;border:1px solid var(--border);background:var(--surface2);color:var(--text-dim);cursor:pointer;font-size:12px;font-weight:600;transition:all .2s;}
.quick-btn:hover{border-color:var(--amber);color:var(--amber);}

.nozzle-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;}
.nozzle-card{background:var(--surface2);border:2px solid var(--border);border-radius:12px;padding:16px;text-align:center;transition:all .2s;position:relative;}
.nozzle-card:hover{border-color:rgba(34,211,238,.5);}
.nozzle-card.offline{opacity:.5;border-color:var(--red);}
.nozzle-id{font-family:'Rajdhani',sans-serif;font-size:32px;font-weight:700;color:var(--petrol);margin-bottom:8px;}
.nozzle-fuel{font-size:12px;color:var(--text-muted);text-transform:uppercase;margin-bottom:6px;letter-spacing:1px;}
.nozzle-status{display:inline-block;padding:4px 12px;border-radius:20px;font-size:10px;font-weight:700;}
.nozzle-status.active{background:var(--green-dim);color:var(--green);}
.nozzle-status.idle{background:var(--amber-dim);color:var(--amber);}
.nozzle-status.offline{background:var(--red-dim);color:var(--red);}
.nozzle-reading{font-family:'DM Mono',monospace;font-size:11px;color:var(--text-dim);margin-top:8px;padding-top:8px;border-top:1px solid var(--border);}
.nozzle-assigned{font-size:11px;color:var(--petrol);margin-top:6px;font-weight:600;}

.btn{padding:9px 18px;border-radius:9px;border:none;cursor:pointer;font-size:12px;font-weight:700;font-family:'Rajdhani',sans-serif;letter-spacing:1.5px;text-transform:uppercase;transition:all .2s;display:inline-flex;align-items:center;gap:7px;white-space:nowrap;}
.btn-primary{background:linear-gradient(135deg,var(--petrol),#0891b2);color:#000;}
.btn-primary:hover{opacity:.9;transform:translateY(-1px);}
.btn-success{background:var(--green-dim);color:var(--green);border:1px solid rgba(16,185,129,.25);}
.btn-success:hover{background:rgba(16,185,129,.2);}
.btn-danger{background:var(--red-dim);color:var(--red);border:1px solid rgba(239,68,68,.25);}
.btn-danger:hover{background:rgba(239,68,68,.2);}
.btn-warning{background:var(--amber-dim);color:var(--amber);border:1px solid rgba(245,158,11,.25);}
.btn-warning:hover{background:rgba(245,158,11,.2);}
.btn-secondary{background:var(--surface2);color:var(--text-dim);border:1px solid var(--border);}
.btn-secondary:hover{border-color:var(--border2);color:var(--text);}
.btn:active{transform:scale(.96)!important;}
.btn-full{width:100%;justify-content:center;}
.btn-sm{padding:5px 11px;font-size:10px;}
.btn-lg{padding:13px 24px;font-size:14px;}

.form-group{margin-bottom:16px;}
.form-label{display:block;font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--text-muted);margin-bottom:7px;font-weight:600;}
.form-input,.form-select{width:100%;background:var(--surface2);border:1px solid var(--border);border-radius:9px;padding:11px 14px;color:var(--text);font-size:14px;font-family:'DM Sans',sans-serif;outline:none;transition:border-color .2s,box-shadow .2s;}
.form-input:focus,.form-select:focus{border-color:var(--petrol);box-shadow:0 0 0 3px var(--petrol-dim);}
.form-input::placeholder{color:var(--text-muted);}
.form-select{cursor:pointer;-webkit-appearance:none;appearance:none;background-image:url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 16 16'%3e%3cpath fill='none' stroke='%2364748b' stroke-linecap='round' stroke-linejoin='round' stroke-width='2' d='M2 5l6 6 6-6'/%3e%3c/svg%3e");background-repeat:no-repeat;background-position:right 12px center;background-size:12px;}
.form-select option{background:#121c2e;}

.table-wrap{overflow-x:auto;}
.data-table{width:100%;border-collapse:collapse;}
.data-table th{font-size:9px;letter-spacing:2px;text-transform:uppercase;color:var(--text-muted);padding:10px 14px;text-align:left;border-bottom:1px solid var(--border);white-space:nowrap;}
.data-table td{padding:10px 14px;border-bottom:1px solid rgba(255,255,255,.03);font-size:12px;color:var(--text-dim);white-space:nowrap;}
.data-table tr:hover td{background:rgba(255,255,255,.02);color:var(--text);}
.badge{display:inline-flex;align-items:center;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;}
.badge-petrol{background:var(--petrol-dim);color:var(--petrol);}
.badge-diesel{background:var(--diesel-dim);color:var(--diesel);}
.badge-green{background:var(--green-dim);color:var(--green);}
.badge-red{background:var(--red-dim);color:var(--red);}
.badge-blue{background:var(--blue-dim);color:var(--blue);}

.notice-board{background:var(--surface2);border:1px solid var(--border);border-left:4px solid var(--amber);border-radius:10px;padding:16px;margin-bottom:12px;}
.notice-time{font-size:10px;color:var(--text-muted);margin-bottom:4px;}
.notice-msg{font-size:13px;color:var(--text);line-height:1.5;}
.notice-by{font-size:10px;color:var(--amber);margin-top:8px;}

.chart-container{position:relative;width:100%;}
.chart-sm{height:200px;}
.chart-md{height:260px;}
.chart-lg{height:320px;}

.notification{position:fixed;top:16px;right:16px;padding:13px 18px;border-radius:12px;font-size:13px;font-weight:600;z-index:9999;transform:translateX(130%);transition:transform .35s cubic-bezier(.34,1.56,.64,1);display:flex;align-items:center;gap:10px;max-width:340px;box-shadow:0 10px 40px rgba(0,0,0,.5);}
.notification.show{transform:translateX(0);}
.notification.success{background:rgba(6,78,59,.95);border:1px solid var(--green);color:var(--green);}
.notification.error{background:rgba(69,10,10,.95);border:1px solid var(--red);color:var(--red);}
.notification.info{background:rgba(12,40,80,.95);border:1px solid var(--blue);color:var(--blue);}
.notification.warn{background:rgba(69,40,3,.95);border:1px solid var(--amber);color:var(--amber);}

.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.9);z-index:500;display:none;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px);}
.modal-overlay.open{display:flex;animation:fadeIn .2s ease;}
.modal{background:var(--surface);border:1px solid var(--border);border-radius:16px;max-width:500px;width:100%;animation:fadeIn .25s ease;max-height:90vh;overflow-y:auto;}
.modal-header{padding:16px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:var(--surface);z-index:10;}
.modal-title{font-family:'Rajdhani',sans-serif;font-size:15px;font-weight:700;letter-spacing:1px;}
.modal-close{width:28px;height:28px;border-radius:6px;border:1px solid var(--border);background:var(--surface2);color:var(--text-muted);cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:16px;line-height:1;}
.modal-close:hover{border-color:var(--red);color:var(--red);}
.modal-body{padding:20px;}

.bill-receipt{background:#fff;color:#111;border-radius:10px;padding:22px;font-family:'DM Sans',sans-serif;}
.bill-station-name{font-family:'Rajdhani',sans-serif;font-size:28px;font-weight:700;letter-spacing:3px;text-align:center;color:#111;}
.bill-location{font-size:10px;color:#9ca3af;text-align:center;margin-bottom:12px;}
.bill-divider{border:none;border-top:1px dashed #d1d5db;margin:10px 0;}
.bill-row{display:flex;justify-content:space-between;padding:4px 0;font-size:12px;}
.bill-lbl{color:#6b7280;}
.bill-val{font-weight:600;color:#111;}
.bill-total-row{background:#111;color:#fff;border-radius:8px;padding:12px 16px;display:flex;justify-content:space-between;align-items:center;margin-top:10px;}
.bill-total-lbl{font-family:'Rajdhani',sans-serif;font-size:13px;font-weight:700;letter-spacing:2px;}
.bill-total-amt{font-family:'Rajdhani',sans-serif;font-size:24px;font-weight:700;color:#10b981;}
.bill-gst-row{font-size:11px;color:#6b7280;padding:4px 0;display:flex;justify-content:space-between;}
.bill-footer{text-align:center;margin-top:12px;font-size:10px;color:#9ca3af;border-top:1px dashed #d1d5db;padding-top:10px;}

.bank-card{background:linear-gradient(135deg,var(--green),var(--blue));border-radius:14px;padding:28px 24px;color:#fff;margin-bottom:20px;box-shadow:0 8px 32px rgba(16,185,129,.3);}
.bank-card-title{font-family:'Rajdhani',sans-serif;font-size:13px;font-weight:600;letter-spacing:2px;text-transform:uppercase;opacity:.8;margin-bottom:16px;}
.bank-card-amount{font-family:'Rajdhani',sans-serif;font-size:48px;font-weight:700;line-height:1;margin-bottom:16px;letter-spacing:-1px;}
.bank-card-details{display:grid;grid-template-columns:1fr 1fr;gap:20px;font-size:12px;}
.bank-card-detail-label{opacity:.7;margin-bottom:4px;text-transform:uppercase;letter-spacing:1px;font-size:9px;}
.bank-card-detail-value{font-family:'DM Mono',monospace;font-size:13px;font-weight:600;}

.shift-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px;margin-bottom:12px;}
.shift-header{display:flex;justify-content:space-between;align-items:center;padding-bottom:12px;border-bottom:1px solid var(--border);}
.shift-info{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-top:12px;font-size:12px;}
.shift-label{color:var(--text-muted);font-size:10px;text-transform:uppercase;letter-spacing:1px;}
.shift-value{font-weight:600;color:var(--text);font-family:'DM Mono',monospace;}

.branch-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:14px;}
.branch-card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:16px;cursor:pointer;transition:all .2s;}
.branch-card:hover{transform:translateY(-2px);border-color:var(--petrol);}
.branch-card.selected{border-color:var(--petrol);background:var(--petrol-dim);}
.branch-name{font-family:'Rajdhani',sans-serif;font-size:14px;font-weight:700;margin-bottom:8px;}
.branch-info{font-size:11px;color:var(--text-muted);margin-bottom:4px;}

.report-tab{padding:8px 16px;border-radius:8px;border:1px solid var(--border);background:var(--surface2);color:var(--text-dim);cursor:pointer;font-size:12px;font-weight:700;font-family:'Rajdhani',sans-serif;letter-spacing:1px;text-transform:uppercase;transition:all .2s;}
.report-tab.active{background:var(--petrol-dim);border-color:rgba(34,211,238,.3);color:var(--petrol);}
.report-tab:hover{border-color:var(--border2);color:var(--text);}

/* Role badge */
.role-tag{display:inline-block;padding:2px 8px;border-radius:10px;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin-left:4px;}
.role-owner{background:var(--amber-dim);color:var(--amber);}
.role-cashier{background:var(--purple-dim);color:var(--purple);}
.role-attendant{background:var(--green-dim);color:var(--green);}

/* Salary row */
.salary-row{display:flex;justify-content:space-between;align-items:center;padding:12px 0;border-bottom:1px solid var(--border);}
.salary-row:last-child{border-bottom:none;}

@media print{
  body *{visibility:hidden!important;}
  #print-area,#print-area *{visibility:visible!important;}
  #print-area{display:block!important;position:fixed!important;top:0;left:0;width:80mm;background:#fff;}
  .header,.nav-wrap,.notification,.btn{display:none!important;}
}
.flex{display:flex;}.gap-8{gap:8px;}.gap-10{gap:10px;}.gap-12{gap:12px;}
.mb-16{margin-bottom:16px;}.mb-20{margin-bottom:20px;}.mt-10{margin-top:10px;}
.text-right{text-align:right;}.ml-auto{margin-left:auto;}
.keyboard-hint{font-size:9px;color:var(--text-muted);display:inline-block;margin-left:8px;padding:4px 8px;background:var(--surface2);border-radius:4px;border:1px solid var(--border);}
.empty-state{text-align:center;padding:40px;color:var(--text-muted);}
.empty-icon{font-size:40px;opacity:.4;margin-bottom:8px;}
.access-denied{background:var(--red-dim);border:1px solid rgba(239,68,68,.3);border-radius:12px;padding:30px;text-align:center;color:var(--red);font-size:14px;font-weight:600;}
</style>
</head>
<body>
<div id="print-area" style="display:none"></div>

<!-- HEADER -->
<div class="header">
  <div class="header-inner">
    <div class="logo-wrap">
      <div class="logo-icon">⛽</div>
      <div class="logo-text">
        <div class="logo-name">VIGNESH FUELS v3.1</div>
        <div class="logo-sub">Complete Management System</div>
      </div>
    </div>
    <div class="header-middle">
      <select class="outlet-selector" id="outlet-selector" onchange="switchOutlet()">
        <option value="">Loading...</option>
      </select>
      <div id="stock-warnings-header"></div>
    </div>
    <div class="user-switcher" id="user-switcher"></div>
    <div class="header-clock">
      <span class="clock-date" id="clock-date">—</span>
      <span><span class="live-dot"></span><span id="clock-time">—</span></span>
    </div>
  </div>
</div>

<!-- NAVIGATION -->
<div class="nav-wrap">
  <div class="nav-tabs" id="nav-tabs"></div>
</div>

<!-- MAIN -->
<div class="main">

  <!-- DASHBOARD -->
  <div class="page active" id="page-dashboard">
    <div class="stats-grid">
      <div class="stat-card green"><div class="stat-icon">₹</div><div class="stat-label">Today's Revenue</div><div class="stat-value" id="d-today-rev">₹0</div><div class="stat-sub" id="d-today-txn">0 transactions</div></div>
      <div class="stat-card blue"><div class="stat-icon">📅</div><div class="stat-label">This Week</div><div class="stat-value" id="d-week-rev">₹0</div><div class="stat-sub">7-day total</div></div>
      <div class="stat-card amber"><div class="stat-icon">📆</div><div class="stat-label">This Month</div><div class="stat-value" id="d-month-rev">₹0</div><div class="stat-sub">30-day total</div></div>
      <div class="stat-card purple"><div class="stat-icon">💰</div><div class="stat-label">Commission Today</div><div class="stat-value" id="d-commission">₹0</div><div class="stat-sub">Petrol+Diesel</div></div>
    </div>
    <div class="fuel-gauges">
      <div class="gauge-card">
        <div class="gauge-title petrol-t">⛽ Petrol Stock</div>
        <div class="gauge-bar-wrap"><div class="gauge-bar petrol-bar" id="petrol-bar" style="width:0%"></div></div>
        <div class="gauge-info"><span class="gauge-liters petrol-t" id="petrol-liters">0 L</span><span class="gauge-pct" id="petrol-pct">0%</span></div>
      </div>
      <div class="gauge-card">
        <div class="gauge-title diesel-t">🛢️ Diesel Stock</div>
        <div class="gauge-bar-wrap"><div class="gauge-bar diesel-bar" id="diesel-bar" style="width:0%"></div></div>
        <div class="gauge-info"><span class="gauge-liters diesel-t" id="diesel-liters">0 L</span><span class="gauge-pct" id="diesel-pct">0%</span></div>
      </div>
    </div>
    <div class="two-col">
      <div class="panel">
        <div class="panel-header"><span class="panel-title">📈 Revenue Trend (7 Days)</span></div>
        <div class="panel-body"><div class="chart-container chart-sm"><canvas id="chart-revenue"></canvas></div></div>
      </div>
      <div class="panel">
        <div class="panel-header"><span class="panel-title">💹 Fuel Rate History</span></div>
        <div class="panel-body"><div class="chart-container chart-sm"><canvas id="chart-rate-history"></canvas></div></div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-header"><span class="panel-title">🕒 Recent Transactions</span><button class="btn btn-sm btn-secondary" onclick="generateDemoData()">⚡ Add Demo Data</button></div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Time</th><th>Type</th><th>Amount (₹)</th><th>GST 18% (₹)</th><th>Liters</th><th>Employee</th></tr></thead>
          <tbody id="recent-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- DISPENSE -->
  <div class="page" id="page-dispense">
    <div class="dispense-wrap">
      <div id="dispense-access-check">
        <!-- Filled dynamically -->
      </div>
    </div>
  </div>

  <!-- NOZZLES -->
  <div class="page" id="page-nozzles">
    <div class="panel">
      <div class="panel-header">
        <span class="panel-title">🔧 Nozzle/Machine Management</span>
        <button class="btn btn-sm btn-primary" onclick="toggleNozzleSetup()">➕ Add Nozzle</button>
      </div>
      <div class="panel-body">
        <div id="nozzle-setup" style="display:none;margin-bottom:20px;padding:16px;background:var(--surface2);border-radius:10px;">
          <div class="three-col" style="margin-bottom:12px;">
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Nozzle ID</label>
              <input type="text" class="form-input" id="new-nozzle-id" placeholder="e.g., N-01">
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Fuel Type</label>
              <select class="form-select" id="new-nozzle-fuel">
                <option value="petrol">⛽ Petrol</option>
                <option value="diesel">🛢️ Diesel</option>
              </select>
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Opening Reading (L)</label>
              <input type="number" class="form-input" id="new-nozzle-reading" placeholder="0" step="0.01">
            </div>
          </div>
          <div class="form-group" style="margin-bottom:12px;">
            <label class="form-label">Assign to Employee</label>
            <select class="form-select" id="new-nozzle-employee"></select>
          </div>
          <button class="btn btn-success" onclick="addNozzle()">✅ Add Nozzle</button>
        </div>
        <div class="nozzle-grid" id="nozzle-grid"></div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📊 Nozzle Readings & Verification</span></div>
      <div class="panel-body">
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr><th>Nozzle ID</th><th>Fuel</th><th>Current Reading</th><th>Assigned To</th><th>Status</th><th>Last Updated</th><th>Action</th></tr></thead>
            <tbody id="nozzle-readings-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- STOCK WARNINGS -->
  <div class="page" id="page-stock-warnings">
    <div class="panel">
      <div class="panel-header"><span class="panel-title">⚠️ Low Stock Warnings & Threshold Settings</span></div>
      <div class="panel-body">
        <div class="three-col" style="margin-bottom:16px;">
          <div class="form-group">
            <label class="form-label">Petrol Warning Threshold (L)</label>
            <input type="number" class="form-input" id="petrol-threshold" value="500" step="100">
          </div>
          <div class="form-group">
            <label class="form-label">Diesel Warning Threshold (L)</label>
            <input type="number" class="form-input" id="diesel-threshold" value="400" step="100">
          </div>
          <div class="form-group">
            <label class="form-label">Critical Low (for urgent order)</label>
            <input type="number" class="form-input" id="critical-threshold" value="200" step="50">
          </div>
        </div>
        <button class="btn btn-primary" onclick="saveThresholds()">💾 Save Thresholds</button>
      </div>
    </div>
    <div class="panel" id="alerts-panel">
      <div class="panel-header"><span class="panel-title">🚨 Active Alerts</span></div>
      <div class="panel-body" id="alerts-body"></div>
    </div>
  </div>

  <!-- SHIFT HANDOVER -->
  <div class="page" id="page-shift-handover">
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📋 Shift Handover Report</span></div>
      <div class="panel-body">
        <div class="three-col" style="margin-bottom:16px;">
          <div class="form-group">
            <label class="form-label">Your Name</label>
            <input type="text" class="form-input" id="handover-name" placeholder="Your name">
          </div>
          <div class="form-group">
            <label class="form-label">Next Shift Employee</label>
            <select class="form-select" id="handover-next"></select>
          </div>
          <div class="form-group">
            <label class="form-label">Shift Type</label>
            <select class="form-select" id="handover-type">
              <option value="morning">🌅 Morning (6AM-2PM)</option>
              <option value="afternoon">☀️ Afternoon (2PM-10PM)</option>
              <option value="night">🌙 Night (10PM-6AM)</option>
            </select>
          </div>
        </div>
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:16px;margin-bottom:16px;">
          <div style="font-size:10px;letter-spacing:1px;text-transform:uppercase;color:var(--petrol);font-weight:700;margin-bottom:12px;">⛽ Petrol — Auto-Loaded from System</div>
          <div class="three-col">
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Opening Petrol Stock (L)</label>
              <input type="number" class="form-input" id="handover-opening-petrol" placeholder="Auto-loaded" step="0.01" readonly style="opacity:.7">
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Current/Closing Petrol Stock (L)</label>
              <input type="number" class="form-input" id="handover-closing-petrol" placeholder="Auto-loaded" step="0.01" readonly style="opacity:.7">
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Petrol Sold This Shift (L)</label>
              <input type="number" class="form-input" id="handover-petrol-sold" placeholder="Auto-calculated" readonly style="color:var(--petrol);font-weight:600">
            </div>
          </div>
        </div>
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:16px;margin-bottom:16px;">
          <div style="font-size:10px;letter-spacing:1px;text-transform:uppercase;color:var(--diesel);font-weight:700;margin-bottom:12px;">🛢️ Diesel — Auto-Loaded from System</div>
          <div class="three-col">
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Opening Diesel Stock (L)</label>
              <input type="number" class="form-input" id="handover-opening-diesel" placeholder="Auto-loaded" step="0.01" readonly style="opacity:.7">
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Current/Closing Diesel Stock (L)</label>
              <input type="number" class="form-input" id="handover-closing-diesel" placeholder="Auto-loaded" step="0.01" readonly style="opacity:.7">
            </div>
            <div class="form-group" style="margin-bottom:0">
              <label class="form-label">Diesel Sold This Shift (L)</label>
              <input type="number" class="form-input" id="handover-diesel-sold" placeholder="Auto-calculated" readonly style="color:var(--diesel);font-weight:600">
            </div>
          </div>
        </div>
        <div class="form-group">
          <label class="form-label">Remarks/Notes</label>
          <textarea class="form-input" id="handover-remarks" placeholder="Any issues, spillage, equipment problems..." style="resize:vertical;min-height:80px;"></textarea>
        </div>
        <button class="btn btn-success btn-lg" onclick="submitHandover()">✅ Submit Shift Handover</button>
      </div>
    </div>
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📊 Recent Handovers</span></div>
      <div class="panel-body" id="handover-history"></div>
    </div>
  </div>

  <!-- BRANCHES -->
  <div class="page" id="page-branches">
    <div class="panel">
      <div class="panel-header"><span class="panel-title">🏪 Multi-Outlet/Branch Management</span></div>
      <div class="panel-body">
        <div id="branch-setup" style="margin-bottom:20px;padding:16px;background:var(--surface2);border-radius:10px;">
          <div class="three-col" style="margin-bottom:12px;">
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Branch Name</label>
              <input type="text" class="form-input" id="new-branch-name" placeholder="e.g., Guntur Main">
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Location/Address</label>
              <input type="text" class="form-input" id="new-branch-location" placeholder="e.g., NH-16, Near Railway">
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Manager Name</label>
              <input type="text" class="form-input" id="new-branch-manager" placeholder="Manager name">
            </div>
          </div>
          <button class="btn btn-success" onclick="addBranch()">✅ Add Branch</button>
        </div>
        <div class="branch-grid" id="branch-grid"></div>
      </div>
    </div>
  </div>

  <!-- BANK ACCOUNT -->
  <div class="page" id="page-bank-account">
    <div id="bank-account-content"></div>
  </div>

  <!-- EMPLOYEES -->
  <div class="page" id="page-employees">
    <div class="panel">
      <div class="panel-header">
        <span class="panel-title">👥 Employee Management</span>
        <button class="btn btn-sm btn-primary" onclick="toggleEmpSetup()">➕ Add Employee</button>
      </div>
      <div class="panel-body">
        <div id="emp-setup" style="display:none;margin-bottom:20px;padding:16px;background:var(--surface2);border-radius:10px;">
          <div class="three-col" style="margin-bottom:12px;">
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Name</label>
              <input type="text" class="form-input" id="new-emp-name" placeholder="Employee name">
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Role</label>
              <select class="form-select" id="new-emp-role">
                <option value="attendant">⛽ Fuel Attendant</option>
                <option value="cashier">🧾 Cashier</option>
              </select>
            </div>
            <div class="form-group" style="margin-bottom:8px;">
              <label class="form-label">Monthly Salary (₹)</label>
              <input type="number" class="form-input" id="new-emp-salary" placeholder="15000">
            </div>
          </div>
          <button class="btn btn-success" onclick="addEmployee()">✅ Add Employee</button>
        </div>
        <div id="emp-list"></div>
      </div>
    </div>
  </div>

  <!-- NOTICES -->
  <div class="page" id="page-notices">
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📢 Announcement/Notice Board</span></div>
      <div class="panel-body">
        <div id="notice-post" style="padding:16px;background:var(--surface2);border-radius:10px;margin-bottom:20px;">
          <div class="form-group">
            <label class="form-label">New Announcement</label>
            <textarea class="form-input" id="new-notice" placeholder="Type your announcement..." style="resize:vertical;min-height:80px;"></textarea>
          </div>
          <button class="btn btn-primary" onclick="postNotice()">📌 Post Announcement</button>
        </div>
        <div id="notices-feed"></div>
      </div>
    </div>
  </div>

  <!-- REPORTS -->
  <div class="page" id="page-reports">
    <div style="display:flex;gap:6px;margin-bottom:20px;flex-wrap:wrap;align-items:center;">
      <button class="report-tab active" onclick="setReportPeriod('today',this)">Today</button>
      <button class="report-tab" onclick="setReportPeriod('week',this)">This Week</button>
      <button class="report-tab" onclick="setReportPeriod('month',this)">This Month</button>
      <button class="report-tab" onclick="setReportPeriod('all',this)">All Time</button>
      <button class="btn btn-sm btn-secondary" style="margin-left:auto" onclick="exportToExcel()">📊 Export Excel</button>
    </div>
    <div class="stats-grid" id="report-stats"></div>
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📋 Transaction History</span></div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Date</th><th>Time</th><th>Fuel</th><th>Amount (₹)</th><th>GST (₹)</th><th>Liters</th><th>Price/L</th><th>Employee</th></tr></thead>
          <tbody id="report-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- SETTINGS -->
  <div class="page" id="page-settings">
    <div class="two-col">
      <div class="panel">
        <div class="panel-header"><span class="panel-title">⛽ Fuel Prices & Rates</span></div>
        <div class="panel-body">
          <div class="form-group">
            <label class="form-label">Petrol Price (₹/L)</label>
            <input type="number" class="form-input" id="set-petrol-price" step="0.01">
          </div>
          <div class="form-group">
            <label class="form-label">Diesel Price (₹/L)</label>
            <input type="number" class="form-input" id="set-diesel-price" step="0.01">
          </div>
          <button class="btn btn-primary btn-full" onclick="updatePrices()">💾 Update Prices</button>
          <div style="margin-top:12px;padding:12px;background:var(--surface2);border-radius:8px;font-size:11px;color:var(--text-muted);">
            <strong>Live API Status:</strong> <span id="api-status">Checking...</span>
          </div>
        </div>
      </div>
      <div class="panel">
        <div class="panel-header"><span class="panel-title">➕ Stock Management</span></div>
        <div class="panel-body">
          <div class="form-group">
            <label class="form-label">Add Petrol (L)</label>
            <input type="number" class="form-input" id="add-petrol">
          </div>
          <div class="form-group">
            <label class="form-label">Add Diesel (L)</label>
            <input type="number" class="form-input" id="add-diesel">
          </div>
          <div class="flex gap-8">
            <button class="btn btn-success" onclick="addStock()">➕ Add Stock</button>
            <button class="btn btn-warning" onclick="resetStock()">🔄 Reset Full</button>
          </div>
        </div>
      </div>
    </div>
    <div class="panel">
      <div class="panel-header"><span class="panel-title">🗑️ Data Management</span></div>
      <div class="panel-body">
        <div class="flex gap-8" style="flex-wrap:wrap;">
          <button class="btn btn-danger" onclick="if(confirm('Clear ALL data for this branch?'))clearAllData()">🗑️ Clear All Data</button>
          <button class="btn btn-secondary" onclick="exportData()">📥 Export JSON</button>
          <button class="btn btn-secondary" onclick="fetchLivePrices()">🌐 Fetch Live Prices</button>
        </div>
      </div>
    </div>
  </div>

</div><!-- /main -->

<!-- BILL MODAL -->
<div class="modal-overlay" id="bill-modal" onclick="if(event.target===this)closeBill()">
  <div class="modal">
    <div class="modal-header">
      <span class="modal-title">🧾 Fuel Bill</span>
      <button class="modal-close" onclick="closeBill()">×</button>
    </div>
    <div class="modal-body">
      <div id="bill-html"></div>
      <div class="flex gap-8 mt-10">
        <button class="btn btn-primary btn-full" onclick="doPrint()">🖨️ Print Bill</button>
        <button class="btn btn-secondary" onclick="closeBill()">Close</button>
      </div>
    </div>
  </div>
</div>

<div class="notification" id="notif"></div>

<script>
'use strict';

/* ═══════════════════════════════════
   DEFAULTS
═══════════════════════════════════ */
const DEFAULTS = {
  petrolPrice: 110.10,
  dieselPrice: 97.92,
  petrolCommission: 4.00,
  dieselCommission: 3.00,
  gstRate: 0.18,
  petrolThreshold: 500,
  dieselThreshold: 400,
  criticalThreshold: 200
};

// Default employees (shared template — each branch gets its own copy on first load)
const DEFAULT_EMPLOYEES = [
  {id:'admin', name:'Vignesh', role:'Owner/Admin', roleKey:'admin', salary:50000, isAdmin:true, isCashier:false, isAttendant:false, present:true},
  {id:'raju',  name:'Rajesh Kumar', role:'Fuel Attendant', roleKey:'attendant', salary:15000, isAdmin:false, isCashier:false, isAttendant:true, present:true},
  {id:'suresh',name:'Suresh Babu',  role:'Fuel Attendant', roleKey:'attendant', salary:15000, isAdmin:false, isCashier:false, isAttendant:true, present:false},
  {id:'priya', name:'Priya Reddy',  role:'Cashier',        roleKey:'cashier',   salary:18000, isAdmin:false, isCashier:true,  isAttendant:false, present:false}
];

const DEFAULT_BRANCHES = [
  {id:'br-01', name:'Guntur Main', location:'NH-16, Near Railway', manager:'Vignesh'}
];

/* ═══════════════════════════════════
   BRANCH-SCOPED STORAGE
   All data is prefixed with branchId
═══════════════════════════════════ */
let currentBranch = null; // set during init

const load = (k,d) => { try{ const v=localStorage.getItem(k); return v?JSON.parse(v):d; }catch(e){ return d; }};
const save = (k,v) => localStorage.setItem(k, JSON.stringify(v));

// Branch-scoped keys
function bk(key) { return `${currentBranch}__${key}`; }

function getPrices()       { return load(bk('prices'), {petrol:DEFAULTS.petrolPrice, diesel:DEFAULTS.dieselPrice}); }
function savePrices(d)     { save(bk('prices'), d); }
function getStock()        { return load(bk('stock'), {petrol:5200, diesel:6600}); }
function saveStock(d)      { save(bk('stock'), d); }
function getTxns()         { return load(bk('txns'), []); }
function saveTxns(d)       { save(bk('txns'), d); }
function getEmployees()    { return load(bk('employees'), DEFAULT_EMPLOYEES); }
function saveEmployees(d)  { save(bk('employees'), d); }
function getNozzles()      { return load(bk('nozzles'), []); }
function saveNozzles(d)    { save(bk('nozzles'), d); }
function getThresholds()   { return load(bk('thresholds'), {petrol:DEFAULTS.petrolThreshold, diesel:DEFAULTS.dieselThreshold, critical:DEFAULTS.criticalThreshold}); }
function saveThresholds_s(d){ save(bk('thresholds'), d); }
function getBank()         { return load(bk('bank'), {totalEarnings:0, balance:0, lastUpdate:getToday(), transactions:[], salaryPayments:[]}); }
function saveBank(d)       { save(bk('bank'), d); }
function getHandovers()    { return load(bk('handovers'), []); }
function saveHandovers(d)  { save(bk('handovers'), d); }
function getNotices()      { return load(bk('notices'), []); }
function saveNotices(d)    { save(bk('notices'), d); }
function getPriceHistory() { return load(bk('priceHistory'), []); }
function savePriceHistory(d){ save(bk('priceHistory'), d); }
// Shift opening stock per shift type per day
function getShiftOpenStock(){ return load(bk('shiftOpenStock'), {}); }
function saveShiftOpenStock(d){ save(bk('shiftOpenStock'), d); }

// Global (not branch-scoped)
function getBranches()     { return load('vfs_branches', DEFAULT_BRANCHES); }
function saveBranches(d)   { save('vfs_branches', d); }

/* ═══════════════════════════════════
   STATE
═══════════════════════════════════ */
let currentUser = {id:'admin', name:'Vignesh', isAdmin:true, isCashier:false, isAttendant:false};
let selectedFuel = 'petrol';
let reportPeriod = 'today';
let charts = {};

/* ═══════════════════════════════════
   UTILITY
═══════════════════════════════════ */
function getToday() { return new Date().toISOString().split('T')[0]; }
function isWithinDays(dateStr, days) { const d=new Date(dateStr), now=new Date(); return (now-d)/86400000<=days; }
function fmt(n) { return n.toLocaleString('en-IN',{minimumFractionDigits:0,maximumFractionDigits:2}); }
function setT(id,val){ const el=document.getElementById(id); if(el)el.textContent=val; }

function playClick()   { beep(440,.08,'sine',.1); }
function playSuccess() { beep(660,.1);setTimeout(()=>beep(880,.12),100); }
function playError()   { beep(220,.3,'sawtooth',.12); }
function playDispense(){ beep(550,.1);setTimeout(()=>beep(770,.1),90);setTimeout(()=>beep(990,.15),180); }

let audioCtx=null;
function initAudio(){if(!audioCtx)try{audioCtx=new(window.AudioContext||window.webkitAudioContext)();}catch(e){}}
function beep(freq,dur,type='sine',vol=0.15){
  try{
    initAudio();if(!audioCtx)return;
    const o=audioCtx.createOscillator(),g=audioCtx.createGain();
    o.connect(g);g.connect(audioCtx.destination);
    o.frequency.value=freq;o.type=type;
    g.gain.setValueAtTime(vol,audioCtx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001,audioCtx.currentTime+dur);
    o.start();o.stop(audioCtx.currentTime+dur);
  }catch(e){}
}

function notify(msg,type='success'){
  const n=document.getElementById('notif');
  const icons={success:'✅',error:'❌',info:'ℹ️',warn:'⚠️'};
  n.className=`notification ${type}`;
  n.innerHTML=`${icons[type]||''} ${msg}`;
  n.classList.add('show');
  clearTimeout(n._t);
  n._t=setTimeout(()=>n.classList.remove('show'),4000);
}

/* ═══════════════════════════════════
   CLOCK
═══════════════════════════════════ */
function startClock(){
  const tick=()=>{
    const now=new Date();
    setT('clock-time',now.toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true}).toUpperCase());
    setT('clock-date',now.toLocaleDateString('en-IN',{weekday:'short',day:'2-digit',month:'short',year:'numeric'}));
  };
  tick();setInterval(tick,1000);
}

/* ═══════════════════════════════════
   ROLE HELPERS
═══════════════════════════════════ */
function canDispense(user) {
  return user.isAttendant && !user.isAdmin && !user.isCashier;
}
function isAdminOrCashier(user) {
  return user.isAdmin || user.isCashier;
}
function getRoleLabel(emp) {
  if(emp.isAdmin) return '<span class="role-tag role-owner">Owner</span>';
  if(emp.isCashier) return '<span class="role-tag role-cashier">Cashier</span>';
  if(emp.isAttendant) return '<span class="role-tag role-attendant">Attendant</span>';
  return '';
}

/* ═══════════════════════════════════
   USER SWITCHER
═══════════════════════════════════ */
function buildUserSwitcher(){
  const emps=getEmployees();
  const wrap=document.getElementById('user-switcher');
  wrap.innerHTML=emps.map(e=>{
    let cls='user-btn';
    if(e.isAdmin) cls+=' admin';
    else if(e.isCashier) cls+=' cashier';
    if(currentUser.id===e.id) cls+=' active';
    return `<button class="${cls}" onclick="switchUser('${e.id}')">${e.isAdmin?'👑':e.isCashier?'🧾':'👤'} ${e.name.split(' ')[0]}</button>`;
  }).join('');
}

function switchUser(id){
  playClick();
  const emps=getEmployees();
  const emp=emps.find(e=>e.id===id);
  if(!emp)return;
  currentUser={id:emp.id, name:emp.name, isAdmin:emp.isAdmin, isCashier:!!emp.isCashier, isAttendant:!!emp.isAttendant};
  buildUserSwitcher();
  buildNavTabs();
  showPage(currentUser.isAdmin||currentUser.isCashier?'dashboard':'dispense');
  notify(`Welcome, ${emp.name}!`,'info');
}

/* ═══════════════════════════════════
   NAV TABS
   Admin & Cashier: all tabs (no dispense for them)
   Attendant: only Dispense, Shift, Notices, Reports
═══════════════════════════════════ */
const ALL_TABS=[
  {id:'dashboard',    label:'Dashboard',    icon:'📊', adminCashier:true,  attendant:false},
  {id:'dispense',     label:'Dispense',     icon:'⛽', adminCashier:false, attendant:true},
  {id:'nozzles',      label:'Nozzles',      icon:'🔧', adminCashier:true,  attendant:false},
  {id:'stock-warnings',label:'Stock Alerts',icon:'⚠️', adminCashier:true,  attendant:false},
  {id:'shift-handover',label:'Shift',       icon:'📋', adminCashier:true,  attendant:true},
  {id:'branches',     label:'Branches',     icon:'🏪', adminCashier:true,  attendant:false},
  {id:'bank-account', label:'Bank Earnings',icon:'💰', adminCashier:true,  attendant:false},
  {id:'employees',    label:'Employees',    icon:'👥', adminCashier:true,  attendant:false},
  {id:'notices',      label:'Notices',      icon:'📢', adminCashier:true,  attendant:true},
  {id:'reports',      label:'Reports',      icon:'📈', adminCashier:true,  attendant:true},
  {id:'settings',     label:'Settings',     icon:'⚙️', adminCashier:true,  attendant:false}
];

function buildNavTabs(){
  const wrap=document.getElementById('nav-tabs');
  const isAC=isAdminOrCashier(currentUser);
  const tabs=ALL_TABS.filter(t=>isAC?t.adminCashier:t.attendant);
  wrap.innerHTML=tabs.map(t=>`
    <div class="nav-tab" id="nav-${t.id}" onclick="showPage('${t.id}')">
      <span class="tab-icon">${t.icon}</span> ${t.label}
    </div>`).join('');
}

function showPage(id){
  playClick();
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(n=>n.classList.remove('active'));
  const pg=document.getElementById('page-'+id);
  if(pg)pg.classList.add('active');
  const nt=document.getElementById('nav-'+id);
  if(nt)nt.classList.add('active');

  if(id==='dashboard')     refreshDashboard();
  if(id==='dispense')      updateDispensePage();
  if(id==='nozzles')       renderNozzles();
  if(id==='stock-warnings')renderStockWarnings();
  if(id==='shift-handover')renderShiftHandover();
  if(id==='branches')      renderBranches();
  if(id==='bank-account')  renderBankAccount();
  if(id==='employees')     renderEmployees();
  if(id==='notices')       renderNotices();
  if(id==='reports')       renderReports();
  if(id==='settings')      loadSettingsInputs();
}

/* ═══════════════════════════════════
   OUTLET / BRANCH SELECTOR
═══════════════════════════════════ */
function buildOutletSelector(){
  const branches=getBranches();
  const sel=document.getElementById('outlet-selector');
  sel.innerHTML=branches.map(b=>`<option value="${b.id}">${b.name}</option>`).join('');
  if(!currentBranch||!branches.find(b=>b.id===currentBranch)){
    currentBranch=branches[0]?.id;
  }
  sel.value=currentBranch;
}

function switchOutlet(){
  playClick();
  currentBranch=document.getElementById('outlet-selector').value;
  // Reload user (in case employees differ per branch)
  const emps=getEmployees();
  const emp=emps.find(e=>e.id===currentUser.id)||emps[0];
  if(emp) currentUser={id:emp.id,name:emp.name,isAdmin:emp.isAdmin,isCashier:!!emp.isCashier,isAttendant:!!emp.isAttendant};
  buildUserSwitcher();
  buildNavTabs();
  showPage(currentUser.isAdmin||currentUser.isCashier?'dashboard':'dispense');
  notify(`Switched to branch: ${getBranches().find(b=>b.id===currentBranch)?.name}`,'info');
}

/* ═══════════════════════════════════
   STOCK WARNINGS HEADER
═══════════════════════════════════ */
function updateStockWarningsHeader(){
  const stock=getStock(), thresh=getThresholds();
  const warns=[];
  if(stock.petrol<=thresh.petrol) warns.push(`⚠️ Petrol: ${stock.petrol.toFixed(0)}L`);
  if(stock.diesel<=thresh.diesel) warns.push(`⚠️ Diesel: ${stock.diesel.toFixed(0)}L`);
  const el=document.getElementById('stock-warnings-header');
  if(el) el.innerHTML=warns.length?`<div class="stock-warning">${warns.join(' | ')}</div>`:'';
}

/* ═══════════════════════════════════
   DASHBOARD
═══════════════════════════════════ */
function refreshDashboard(){
  const txns=getTxns(), today=getToday(), stock=getStock();
  const todayT=txns.filter(t=>t.date===today);
  const weekT=txns.filter(t=>isWithinDays(t.date,7));
  const monthT=txns.filter(t=>isWithinDays(t.date,30));
  const todayRev=todayT.reduce((s,t)=>s+t.amount,0);
  const weekRev=weekT.reduce((s,t)=>s+t.amount,0);
  const monthRev=monthT.reduce((s,t)=>s+t.amount,0);
  const todayComm=todayT.reduce((s,t)=>s+(t.fuel==='petrol'?t.liters*4:t.liters*3),0);
  setT('d-today-rev',`₹${fmt(todayRev)}`);
  setT('d-today-txn',`${todayT.length} transactions`);
  setT('d-week-rev',`₹${fmt(weekRev)}`);
  setT('d-month-rev',`₹${fmt(monthRev)}`);
  setT('d-commission',`₹${fmt(todayComm)}`);
  const pPct=Math.min(100,(stock.petrol/8000)*100);
  const dPct=Math.min(100,(stock.diesel/12000)*100);
  document.getElementById('petrol-bar').style.width=pPct+'%';
  document.getElementById('diesel-bar').style.width=dPct+'%';
  setT('petrol-liters',`${stock.petrol.toFixed(0)} L`);
  setT('diesel-liters',`${stock.diesel.toFixed(0)} L`);
  setT('petrol-pct',`${pPct.toFixed(0)}%`);
  setT('diesel-pct',`${dPct.toFixed(0)}%`);
  updateStockWarningsHeader();
  buildRevenueChart();
  buildRateHistoryChart();
  renderRecentTable();
}

function renderRecentTable(){
  const txns=getTxns().slice(-20).reverse();
  const tbody=document.getElementById('recent-tbody');
  if(!txns.length){tbody.innerHTML=`<tr><td colspan="6"><div class="empty-state">No transactions yet</div></td></tr>`;return;}
  tbody.innerHTML=txns.map(t=>{
    const gst=t.amount*0.18;
    return `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px">${t.time}</td>
      <td><span class="badge badge-${t.fuel}">${t.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</span></td>
      <td style="color:var(--green);font-weight:600">₹${fmt(t.amount)}</td>
      <td style="color:var(--blue)">₹${fmt(gst)}</td>
      <td style="font-family:'DM Mono',monospace">${t.liters.toFixed(2)} L</td>
      <td>${t.employee}</td>
    </tr>`;
  }).join('');
}

/* ═══════════════════════════════════
   DISPENSE — Only Attendants can dispense
═══════════════════════════════════ */
function updateDispensePage(){
  const wrap=document.getElementById('dispense-access-check');
  if(!canDispense(currentUser)){
    wrap.innerHTML=`<div class="access-denied">
      ⛔ <strong>Dispensing Not Allowed</strong><br><br>
      Only Fuel Attendants can dispense fuel.<br>
      Admins and Cashiers manage the system but do not operate the dispensing pump.<br><br>
      <span style="font-size:12px;opacity:.7">Switch to an Attendant account to dispense.</span>
    </div>`;
    return;
  }
  const prices=getPrices(), stock=getStock();
  wrap.innerHTML=`
    <div class="panel">
      <div class="panel-header"><span class="panel-title">⛽ Fuel Dispensing</span><span class="keyboard-hint">ENTER to dispense | ESC to close</span></div>
      <div class="panel-body">
        <div class="fuel-type-select">
          <div class="fuel-type-btn sel-petrol" id="ftype-petrol" onclick="selectFuel('petrol')">
            <span class="ftype-icon">⛽</span>
            <div class="ftype-name">PETROL</div>
            <div class="ftype-price" id="petrol-price-display">₹${prices.petrol.toFixed(2)} / L</div>
            <div class="ftype-price">₹4.00 commission/L</div>
          </div>
          <div class="fuel-type-btn" id="ftype-diesel" onclick="selectFuel('diesel')">
            <span class="ftype-icon">🛢️</span>
            <div class="ftype-name">DIESEL</div>
            <div class="ftype-price" id="diesel-price-display">₹${prices.diesel.toFixed(2)} / L</div>
            <div class="ftype-price">₹3.00 commission/L</div>
          </div>
        </div>
        <div class="form-label" style="margin-bottom:4px">Quick Amounts</div>
        <div class="quick-amounts">
          <button class="quick-btn" onclick="setAmount(100)">₹100</button>
          <button class="quick-btn" onclick="setAmount(200)">₹200</button>
          <button class="quick-btn" onclick="setAmount(500)">₹500</button>
          <button class="quick-btn" onclick="setAmount(1000)">₹1000</button>
          <button class="quick-btn" onclick="setAmount(2000)">₹2000</button>
          <button class="quick-btn" onclick="setAmount(5000)">₹5000</button>
        </div>
        <div class="form-label" style="margin-bottom:7px;margin-top:14px">Enter Customer Payment Amount (₹)</div>
        <div class="amount-input-wrap">
          <span class="amount-prefix">₹</span>
          <input type="number" class="amount-input" id="dispense-amount" placeholder="0" oninput="calcLiters()" onkeydown="if(event.key==='Enter')doDispense()">
        </div>
        <div class="liters-preview">
          <span class="liters-label">Liters to Dispense</span>
          <span class="liters-val" id="liters-preview">0.00 L</span>
        </div>
        <button class="dispense-btn" onclick="doDispense()">
          <span>⛽</span> DISPENSE FUEL (ENTER)
        </button>
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:14px;margin-top:16px;font-size:12px;">
          <div style="display:flex;justify-content:space-between;padding:4px 0;"><span style="color:var(--text-muted)">⛽ Petrol Stock</span><span style="font-family:'DM Mono',monospace;color:var(--petrol)">${stock.petrol.toFixed(0)} L</span></div>
          <div style="display:flex;justify-content:space-between;padding:4px 0;"><span style="color:var(--text-muted)">🛢️ Diesel Stock</span><span style="font-family:'DM Mono',monospace;color:var(--diesel)">${stock.diesel.toFixed(0)} L</span></div>
          <div style="display:flex;justify-content:space-between;padding:4px 0;"><span style="color:var(--text-muted)">Dispensing as</span><span style="color:var(--green);font-weight:600">${currentUser.name}</span></div>
        </div>
      </div>
    </div>`;
  selectedFuel='petrol';
  setTimeout(()=>{const a=document.getElementById('dispense-amount');if(a)a.focus();},100);
}

function selectFuel(type){
  playClick();
  selectedFuel=type;
  const bp=document.getElementById('ftype-petrol');
  const bd=document.getElementById('ftype-diesel');
  if(bp) bp.className=`fuel-type-btn${type==='petrol'?' sel-petrol':''}`;
  if(bd) bd.className=`fuel-type-btn${type==='diesel'?' sel-diesel':''}`;
  calcLiters();
}

function setAmount(amt){
  playClick();
  const a=document.getElementById('dispense-amount');
  if(a){a.value=amt;calcLiters();}
}

function calcLiters(){
  const amt=parseFloat(document.getElementById('dispense-amount')?.value)||0;
  const price=selectedFuel==='petrol'?getPrices().petrol:getPrices().diesel;
  setT('liters-preview',`${(price>0?amt/price:0).toFixed(2)} L`);
}

function doDispense(){
  initAudio();
  if(!canDispense(currentUser)){playError();notify('Only Attendants can dispense!','error');return;}
  const amtEl=document.getElementById('dispense-amount');
  const amt=parseFloat(amtEl?.value);
  if(!amt||amt<=0){playError();notify('Enter valid amount!','error');return;}
  const prices=getPrices();
  const price=selectedFuel==='petrol'?prices.petrol:prices.diesel;
  const liters=amt/price;
  const stock=getStock();
  if(selectedFuel==='petrol'&&stock.petrol<liters){playError();notify(`Insufficient Petrol! Only ${stock.petrol.toFixed(1)}L available.`,'error');return;}
  if(selectedFuel==='diesel'&&stock.diesel<liters){playError();notify(`Insufficient Diesel! Only ${stock.diesel.toFixed(1)}L available.`,'error');return;}

  if(selectedFuel==='petrol') stock.petrol-=liters;
  else stock.diesel-=liters;
  saveStock(stock);

  const now=new Date();
  const txn={
    date:getToday(),
    time:now.toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:true}).toUpperCase(),
    fuel:selectedFuel, amount:amt, liters:liters, price:price,
    employee:currentUser.name, branch:currentBranch
  };
  const txns=getTxns(); txns.push(txn); saveTxns(txns);

  const commission=selectedFuel==='petrol'?liters*4:liters*3;
  const bank=getBank();
  bank.totalEarnings=(bank.totalEarnings||0)+commission;
  bank.balance=(bank.balance||0)+commission;
  bank.transactions=bank.transactions||[];
  bank.transactions.push({date:getToday(),time:txn.time,fuel:selectedFuel,liters:liters,commission:commission});
  saveBank(bank);

  playDispense();
  notify(`✅ ${liters.toFixed(2)}L ${selectedFuel} for ₹${fmt(amt)} | Commission: ₹${fmt(commission)}`,'success');
  showBill(txn);
  if(amtEl)amtEl.value='';
  setT('liters-preview','0.00 L');
  updateDispensePage();
  refreshDashboard();
}

/* ═══════════════════════════════════
   BILL
═══════════════════════════════════ */
function showBill(txn){
  const html=buildBillHTML(txn);
  document.getElementById('bill-html').innerHTML=html;
  document.getElementById('print-area').innerHTML=html;
  document.getElementById('bill-modal').classList.add('open');
}

function buildBillHTML(txn){
  const gst=txn.amount*0.18;
  const commission=txn.fuel==='petrol'?txn.liters*4:txn.liters*3;
  return `<div class="bill-receipt">
    <div class="bill-station-name">VIGNESH FUELS</div>
    <div class="bill-location">NH-16, Near Railway Station, Guntur, A.P.</div>
    <hr class="bill-divider">
    <div class="bill-row"><span class="bill-lbl">Date</span><span class="bill-val">${txn.date}</span></div>
    <div class="bill-row"><span class="bill-lbl">Time</span><span class="bill-val">${txn.time}</span></div>
    <div class="bill-row"><span class="bill-lbl">Fuel Type</span><span class="bill-val">${txn.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</span></div>
    <div class="bill-row"><span class="bill-lbl">Rate</span><span class="bill-val">₹${txn.price.toFixed(2)} / Liter</span></div>
    <div class="bill-row"><span class="bill-lbl">Quantity</span><span class="bill-val">${txn.liters.toFixed(2)} Liters</span></div>
    <div class="bill-row"><span class="bill-lbl">Served by</span><span class="bill-val">${txn.employee}</span></div>
    <hr class="bill-divider">
    <div class="bill-row"><span class="bill-lbl">Sub Total</span><span class="bill-val">₹${fmt(txn.amount)}</span></div>
    <div class="bill-gst-row"><span class="bill-lbl">GST 18%</span><span class="bill-val">₹${fmt(gst)}</span></div>
    <div class="bill-total-row">
      <span class="bill-total-lbl">TOTAL AMOUNT</span>
      <span class="bill-total-amt">₹${fmt(txn.amount+gst)}</span>
    </div>
    <div class="bill-footer">Thank you for visiting Vignesh Fuel Station!<br>Have a safe journey 🙏<br>Commission Earned: ₹${fmt(commission)}</div>
  </div>`;
}

function closeBill(){ document.getElementById('bill-modal').classList.remove('open'); }
function doPrint(){ window.print(); }

document.addEventListener('keydown',e=>{
  if(e.key==='Escape'&&document.getElementById('bill-modal').classList.contains('open')){ playClick();closeBill(); }
});

/* ═══════════════════════════════════
   NOZZLES — Each nozzle assigned to one employee
=══════════════════════════════════ */
function renderNozzles(){
  const nozzles=getNozzles();
  const grid=document.getElementById('nozzle-grid');
  
  // Populate employee dropdown in add form
  const empSel=document.getElementById('new-nozzle-employee');
  if(empSel){
    const attendants=getEmployees().filter(e=>e.isAttendant||e.isAdmin);
    empSel.innerHTML=attendants.map(e=>`<option value="${e.id}">${e.name}</option>`).join('');
  }

  if(!nozzles.length){
    grid.innerHTML='<div class="empty-state" style="grid-column:1/-1"><div class="empty-icon">🔧</div><p>No nozzles added yet</p></div>';
    renderNozzleReadings([]);
    return;
  }
  const emps=getEmployees();
  grid.innerHTML=nozzles.map(n=>{
    const assignedEmp=emps.find(e=>e.id===n.assignedTo);
    return `<div class="nozzle-card ${n.status||'idle'}">
      <div class="nozzle-id">${n.id}</div>
      <div class="nozzle-fuel">${n.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</div>
      <div class="nozzle-status ${n.status||'idle'}">${n.status==='active'?'🟢 Active':(n.status==='offline'?'🔴 Offline':'⚪ Idle')}</div>
      <div class="nozzle-assigned">👤 ${assignedEmp?assignedEmp.name:'Unassigned'}</div>
      <div class="nozzle-reading">Reading: ${n.currentReading||0} L</div>
    </div>`;
  }).join('');
  renderNozzleReadings(nozzles);
}

function toggleNozzleSetup(){
  playClick();
  const s=document.getElementById('nozzle-setup');
  s.style.display=s.style.display==='none'?'block':'none';
}

function addNozzle(){
  const id=document.getElementById('new-nozzle-id').value.trim();
  const fuel=document.getElementById('new-nozzle-fuel').value;
  const reading=parseFloat(document.getElementById('new-nozzle-reading').value)||0;
  const assignedTo=document.getElementById('new-nozzle-employee').value;
  if(!id){playError();notify('Enter Nozzle ID!','error');return;}
  const nozzles=getNozzles();
  if(nozzles.find(n=>n.id===id)){playError();notify('Nozzle ID already exists!','error');return;}
  nozzles.push({id,fuel,currentReading:reading,openingReading:reading,assignedTo:assignedTo,status:'idle',addedDate:getToday()});
  saveNozzles(nozzles);
  document.getElementById('new-nozzle-id').value='';
  document.getElementById('new-nozzle-reading').value='';
  playSuccess();
  notify(`✅ Nozzle ${id} added!`,'success');
  renderNozzles();
}

function renderNozzleReadings(nozzles){
  const tbody=document.getElementById('nozzle-readings-tbody');
  const emps=getEmployees();
  if(!nozzles.length){tbody.innerHTML='<tr><td colspan="7"><div class="empty-state">No nozzles</div></td></tr>';return;}
  tbody.innerHTML=nozzles.map(n=>{
    const assignedEmp=emps.find(e=>e.id===n.assignedTo);
    return `<tr>
      <td><strong>${n.id}</strong></td>
      <td><span class="badge badge-${n.fuel}">${n.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</span></td>
      <td style="font-family:'DM Mono',monospace">${n.currentReading||0} L</td>
      <td>${assignedEmp?`<span style="color:var(--green)">${assignedEmp.name}</span>`:'<span style="color:var(--text-muted)">Unassigned</span>'}</td>
      <td><span class="badge badge-${n.status==='active'?'green':n.status==='offline'?'red':'blue'}">${n.status==='active'?'Active':n.status==='offline'?'Offline':'Idle'}</span></td>
      <td style="font-size:11px">${n.lastUpdated||'Never'}</td>
      <td>
        <button class="btn btn-sm btn-secondary" onclick="updateNozzleReading('${n.id}')">Update</button>
        <button class="btn btn-sm btn-warning" onclick="reassignNozzle('${n.id}')">Reassign</button>
      </td>
    </tr>`;
  }).join('');
}

function updateNozzleReading(id){
  const nozzles=getNozzles(), nozzle=nozzles.find(n=>n.id===id);
  if(!nozzle)return;
  const v=prompt(`Update reading for ${id} (Current: ${nozzle.currentReading}L):`,nozzle.currentReading);
  if(v===null)return;
  nozzle.currentReading=parseFloat(v)||nozzle.currentReading;
  nozzle.lastUpdated=new Date().toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true});
  saveNozzles(nozzles);playSuccess();notify('Nozzle reading updated!','success');renderNozzles();
}

function reassignNozzle(id){
  const emps=getEmployees().filter(e=>e.isAttendant||e.isAdmin);
  const empList=emps.map((e,i)=>`${i+1}. ${e.name}`).join('\n');
  const choice=prompt(`Reassign Nozzle ${id} to:\n${empList}\n\nEnter number:`);
  if(!choice)return;
  const idx=parseInt(choice)-1;
  if(idx<0||idx>=emps.length){notify('Invalid choice','error');return;}
  const nozzles=getNozzles(), nozzle=nozzles.find(n=>n.id===id);
  if(!nozzle)return;
  nozzle.assignedTo=emps[idx].id;
  saveNozzles(nozzles);playSuccess();notify(`Nozzle reassigned to ${emps[idx].name}`,'success');renderNozzles();
}

/* ═══════════════════════════════════
   STOCK WARNINGS
═══════════════════════════════════ */
function renderStockWarnings(){
  const stock=getStock(), thresh=getThresholds();
  document.getElementById('petrol-threshold').value=thresh.petrol;
  document.getElementById('diesel-threshold').value=thresh.diesel;
  document.getElementById('critical-threshold').value=thresh.critical;
  const alerts=[];
  if(stock.petrol<=thresh.critical) alerts.push({type:'critical',fuel:'Petrol',stock:stock.petrol,threshold:thresh.critical});
  if(stock.diesel<=thresh.critical) alerts.push({type:'critical',fuel:'Diesel',stock:stock.diesel,threshold:thresh.critical});
  if(stock.petrol>thresh.critical&&stock.petrol<=thresh.petrol) alerts.push({type:'warning',fuel:'Petrol',stock:stock.petrol,threshold:thresh.petrol});
  if(stock.diesel>thresh.critical&&stock.diesel<=thresh.diesel) alerts.push({type:'warning',fuel:'Diesel',stock:stock.diesel,threshold:thresh.diesel});
  const alertBody=document.getElementById('alerts-body');
  if(!alerts.length){alertBody.innerHTML='<div class="empty-state"><div class="empty-icon">✅</div><p>All stock levels normal</p></div>';return;}
  alertBody.innerHTML=alerts.map(a=>`
    <div style="padding:12px;border-radius:8px;margin-bottom:8px;border-left:4px solid ${a.type==='critical'?'var(--red)':'var(--amber)'};background:${a.type==='critical'?'var(--red-dim)':'var(--amber-dim)'}">
      <div style="font-weight:600;color:${a.type==='critical'?'var(--red)':'var(--amber)'}">${a.type==='critical'?'🚨 CRITICAL':'⚠️ WARNING'}: ${a.fuel}</div>
      <div style="font-size:12px;margin-top:4px">Current: ${a.stock.toFixed(0)}L | Threshold: ${a.threshold}L</div>
    </div>`).join('');
}

function saveThresholds(){
  const thresh={
    petrol:parseFloat(document.getElementById('petrol-threshold').value)||500,
    diesel:parseFloat(document.getElementById('diesel-threshold').value)||400,
    critical:parseFloat(document.getElementById('critical-threshold').value)||200
  };
  saveThresholds_s(thresh);playSuccess();notify('Thresholds saved!','success');renderStockWarnings();
}

/* ═══════════════════════════════════
   SHIFT HANDOVER
   Auto-populate opening/closing from system stock
   Opening = saved when shift was started (or stock at page load)
   Closing = current actual stock
═══════════════════════════════════ */
function renderShiftHandover(){
  const emps=getEmployees();
  const nextSel=document.getElementById('handover-next');
  nextSel.innerHTML=emps.map(e=>`<option value="${e.id}">${e.name}</option>`).join('');
  
  // Auto-populate stock data
  autoFillShiftStock();
  renderHandoverHistory();
}

function autoFillShiftStock(){
  const stock=getStock();
  const shiftType=document.getElementById('handover-type')?.value||'morning';
  const key=`${getToday()}_${shiftType}`;
  const openData=getShiftOpenStock();

  // If no opening stock recorded for this shift yet, record current as opening
  if(!openData[key]){
    openData[key]={petrol:stock.petrol, diesel:stock.diesel};
    saveShiftOpenStock(openData);
  }

  const opening=openData[key];
  const opP=document.getElementById('handover-opening-petrol');
  const clP=document.getElementById('handover-closing-petrol');
  const slP=document.getElementById('handover-petrol-sold');
  const opD=document.getElementById('handover-opening-diesel');
  const clD=document.getElementById('handover-closing-diesel');
  const slD=document.getElementById('handover-diesel-sold');

  if(opP) opP.value=opening.petrol.toFixed(2);
  if(clP) clP.value=stock.petrol.toFixed(2);
  if(slP) slP.value=Math.max(0,opening.petrol-stock.petrol).toFixed(2);
  if(opD) opD.value=opening.diesel.toFixed(2);
  if(clD) clD.value=stock.diesel.toFixed(2);
  if(slD) slD.value=Math.max(0,opening.diesel-stock.diesel).toFixed(2);
}

// When shift type changes, reload the opening stock for that shift
document.addEventListener('change',e=>{
  if(e.target&&e.target.id==='handover-type') autoFillShiftStock();
});

function submitHandover(){
  const name=document.getElementById('handover-name').value.trim();
  const nextId=document.getElementById('handover-next').value;
  const type=document.getElementById('handover-type').value;
  const openP=parseFloat(document.getElementById('handover-opening-petrol').value)||0;
  const closeP=parseFloat(document.getElementById('handover-closing-petrol').value)||0;
  const openD=parseFloat(document.getElementById('handover-opening-diesel').value)||0;
  const closeD=parseFloat(document.getElementById('handover-closing-diesel').value)||0;
  const remarks=document.getElementById('handover-remarks').value.trim();
  if(!name){playError();notify('Enter your name!','error');return;}

  const handovers=getHandovers();
  const nextEmp=getEmployees().find(e=>e.id===nextId);
  const stock=getStock();
  
  // Record next shift's opening stock
  const nextShiftKey=`${getToday()}_${nextEmp?.preferredShift||type}`;
  const openData=getShiftOpenStock();
  openData[`${getToday()}_next_after_${type}`]={petrol:stock.petrol, diesel:stock.diesel};
  saveShiftOpenStock(openData);

  handovers.push({
    date:getToday(),
    time:new Date().toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}),
    fromEmployee:name, toEmployee:nextEmp?.name||'Unknown', shiftType:type,
    openingPetrol:openP, closingPetrol:closeP,
    petrolSold:(openP-closeP).toFixed(2),
    openingDiesel:openD, closingDiesel:closeD,
    dieselSold:(openD-closeD).toFixed(2),
    remarks:remarks
  });
  saveHandovers(handovers);

  document.getElementById('handover-name').value='';
  document.getElementById('handover-remarks').value='';
  playSuccess();notify('✅ Shift handover submitted!','success');
  renderHandoverHistory();
}

function renderHandoverHistory(){
  const handovers=getHandovers().slice(-10).reverse();
  const hist=document.getElementById('handover-history');
  if(!handovers.length){hist.innerHTML='<div class="empty-state">No handovers yet</div>';return;}
  hist.innerHTML=handovers.map(h=>`
    <div class="shift-card">
      <div class="shift-header">
        <div>
          <strong>${h.fromEmployee}</strong> → <strong>${h.toEmployee}</strong>
          <div style="font-size:11px;color:var(--text-muted)">${h.date} ${h.time}</div>
        </div>
        <span class="badge badge-blue">${h.shiftType}</span>
      </div>
      <div class="shift-info">
        <div>
          <div class="shift-label">Petrol Sold</div>
          <div class="shift-value" style="color:var(--petrol)">${Math.max(0,parseFloat(h.petrolSold)||0).toFixed(2)} L</div>
        </div>
        <div>
          <div class="shift-label">Diesel Sold</div>
          <div class="shift-value" style="color:var(--diesel)">${Math.max(0,parseFloat(h.dieselSold)||0).toFixed(2)} L</div>
        </div>
        <div>
          <div class="shift-label">Remarks</div>
          <div class="shift-value" style="font-family:'DM Sans',sans-serif;font-size:11px">${h.remarks||'—'}</div>
        </div>
      </div>
    </div>`).join('');
}

/* ═══════════════════════════════════
   BRANCHES
═══════════════════════════════════ */
function renderBranches(){
  const branches=getBranches();
  const grid=document.getElementById('branch-grid');
  grid.innerHTML=branches.map(b=>`
    <div class="branch-card ${b.id===currentBranch?'selected':''}" onclick="selectBranch('${b.id}')">
      <div class="branch-name">🏪 ${b.name}</div>
      <div class="branch-info">📍 ${b.location}</div>
      <div class="branch-info">👤 Manager: ${b.manager}</div>
      ${branches.length>1?`<button class="btn btn-sm btn-danger" style="width:100%;margin-top:8px" onclick="event.stopPropagation();deleteBranch('${b.id}')">🗑️ Delete</button>`:''}
    </div>`).join('');
}

function selectBranch(id){
  currentBranch=id;
  document.getElementById('outlet-selector').value=id;
  buildUserSwitcher();
  renderBranches();
  refreshDashboard();
  notify('Branch switched!','info');
}

function addBranch(){
  const name=document.getElementById('new-branch-name').value.trim();
  const location=document.getElementById('new-branch-location').value.trim();
  const manager=document.getElementById('new-branch-manager').value.trim();
  if(!name){playError();notify('Enter branch name!','error');return;}
  const branches=getBranches();
  branches.push({id:`br-${Date.now()}`,name,location,manager});
  saveBranches(branches);
  buildOutletSelector();
  document.getElementById('new-branch-name').value='';
  document.getElementById('new-branch-location').value='';
  document.getElementById('new-branch-manager').value='';
  playSuccess();notify(`✅ Branch "${name}" added!`,'success');
  renderBranches();
}

function deleteBranch(id){
  if(!confirm('Delete this branch?'))return;
  const branches=getBranches().filter(b=>b.id!==id);
  if(!branches.length){notify('Cannot delete last branch!','error');return;}
  saveBranches(branches);
  if(currentBranch===id) currentBranch=branches[0].id;
  buildOutletSelector();playSuccess();notify('Branch deleted!','warn');renderBranches();
}

/* ═══════════════════════════════════
   BANK ACCOUNT + PAY SALARIES
   Pay Salaries deducts from bank balance
═══════════════════════════════════ */
function renderBankAccount(){
  if(!isAdminOrCashier(currentUser)){
    document.getElementById('bank-account-content').innerHTML='<div style="padding:40px;text-align:center;color:var(--text-muted)"><strong>⛔ Admin/Cashier Only</strong></div>';
    return;
  }
  const bank=getBank();
  const emps=getEmployees();
  const totalSalary=emps.filter(e=>!e.isAdmin).reduce((s,e)=>s+(e.salary||0),0);

  document.getElementById('bank-account-content').innerHTML=`
    <div class="bank-card">
      <div class="bank-card-title">💰 Commission Earnings — Current Balance</div>
      <div class="bank-card-amount">₹${fmt(bank.balance||0)}</div>
      <div class="bank-card-details">
        <div>
          <div class="bank-card-detail-label">Total Earned (all time)</div>
          <div class="bank-card-detail-value">₹${fmt(bank.totalEarnings||0)}</div>
        </div>
        <div>
          <div class="bank-card-detail-label">Total Salaries Paid</div>
          <div class="bank-card-detail-value">₹${fmt(bank.totalSalariesPaid||0)}</div>
        </div>
      </div>
    </div>

    <div class="panel" style="margin-bottom:20px">
      <div class="panel-header"><span class="panel-title">💸 Pay Salaries — Deducts from Balance</span></div>
      <div class="panel-body">
        <div style="background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:16px;margin-bottom:14px;">
          ${emps.filter(e=>!e.isAdmin).map(e=>`
          <div class="salary-row">
            <div>
              <div style="font-weight:600;font-size:13px">${e.name} ${getRoleLabel(e)}</div>
              <div style="font-size:11px;color:var(--text-muted)">${e.role}</div>
            </div>
            <div style="font-family:'Rajdhani',sans-serif;font-size:18px;color:var(--green)">₹${fmt(e.salary||0)}</div>
          </div>`).join('')}
          <div style="margin-top:12px;padding-top:12px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;">
            <span style="font-weight:700;font-size:13px">Total Monthly Salaries</span>
            <span style="font-family:'Rajdhani',sans-serif;font-size:22px;color:var(--amber)">₹${fmt(totalSalary)}</span>
          </div>
        </div>
        <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
          <button class="btn btn-danger btn-lg" onclick="paySalaries()">💸 Pay All Salaries (₹${fmt(totalSalary)})</button>
          <span style="font-size:11px;color:var(--text-muted)">Balance after payment: ₹${fmt((bank.balance||0)-totalSalary)}</span>
        </div>
      </div>
    </div>

    <div class="panel">
      <div class="panel-header"><span class="panel-title">📊 Commission Transactions</span></div>
      <div class="panel-body">
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr><th>Date</th><th>Time</th><th>Fuel</th><th>Liters</th><th>Rate</th><th>Earned (₹)</th></tr></thead>
            <tbody>${(bank.transactions||[]).slice(-50).reverse().map(t=>`<tr>
              <td style="font-size:11px">${t.date}</td>
              <td style="font-size:11px">${t.time}</td>
              <td><span class="badge badge-${t.fuel}">${t.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</span></td>
              <td style="font-family:'DM Mono',monospace">${t.liters.toFixed(2)} L</td>
              <td>₹${t.fuel==='petrol'?'4.00':'3.00'}/L</td>
              <td style="color:var(--green);font-weight:600">₹${fmt(t.commission)}</td>
            </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>

    ${(bank.salaryPayments||[]).length?`
    <div class="panel">
      <div class="panel-header"><span class="panel-title">📋 Salary Payment History</span></div>
      <div class="panel-body">
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr><th>Date</th><th>Time</th><th>Total Paid</th><th>Balance Before</th><th>Balance After</th></tr></thead>
            <tbody>${(bank.salaryPayments||[]).slice(-20).reverse().map(p=>`<tr>
              <td>${p.date}</td><td>${p.time}</td>
              <td style="color:var(--red)">₹${fmt(p.amount)}</td>
              <td style="font-family:'DM Mono',monospace">₹${fmt(p.balanceBefore)}</td>
              <td style="font-family:'DM Mono',monospace;color:var(--amber)">₹${fmt(p.balanceAfter)}</td>
            </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>`:''}
  `;
}

function paySalaries(){
  const bank=getBank();
  const emps=getEmployees();
  const totalSalary=emps.filter(e=>!e.isAdmin).reduce((s,e)=>s+(e.salary||0),0);
  if(totalSalary<=0){notify('No salaries to pay!','warn');return;}
  const currentBalance=bank.balance||0;
  if(currentBalance<totalSalary){
    if(!confirm(`Warning: Balance (₹${fmt(currentBalance)}) is less than total salaries (₹${fmt(totalSalary)}). Balance will go negative. Continue?`)) return;
  } else {
    if(!confirm(`Pay ₹${fmt(totalSalary)} in salaries? Your balance will decrease from ₹${fmt(currentBalance)} to ₹${fmt(currentBalance-totalSalary)}.`)) return;
  }
  const balanceBefore=currentBalance;
  bank.balance=currentBalance-totalSalary;
  bank.totalSalariesPaid=(bank.totalSalariesPaid||0)+totalSalary;
  bank.salaryPayments=bank.salaryPayments||[];
  bank.salaryPayments.push({
    date:getToday(),
    time:new Date().toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}),
    amount:totalSalary, balanceBefore:balanceBefore, balanceAfter:bank.balance
  });
  saveBank(bank);
  playSuccess();
  notify(`✅ Salaries paid! ₹${fmt(totalSalary)} deducted. New balance: ₹${fmt(bank.balance)}`,'success');
  renderBankAccount();
}

/* ═══════════════════════════════════
   EMPLOYEES — Branch-scoped
═══════════════════════════════════ */
function renderEmployees(){
  const emps=getEmployees();
  const el=document.getElementById('emp-list');
  if(!emps.length){el.innerHTML='<div class="empty-state">No employees</div>';return;}
  el.innerHTML=`<div class="table-wrap"><table class="data-table">
    <thead><tr><th>Name</th><th>Role</th><th>Salary (₹/mo)</th><th>Can Dispense</th><th>Action</th></tr></thead>
    <tbody>${emps.map(e=>`<tr>
      <td><strong>${e.name}</strong> ${getRoleLabel(e)}</td>
      <td style="color:var(--text-dim)">${e.role}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--green)">₹${fmt(e.salary||0)}</td>
      <td>${e.isAttendant?'<span class="badge badge-green">✅ Yes</span>':'<span class="badge badge-red">⛔ No</span>'}</td>
      <td>${!e.isAdmin?`<button class="btn btn-sm btn-danger" onclick="deleteEmployee('${e.id}')">🗑️</button>`:'<span style="color:var(--text-muted);font-size:11px">Owner</span>'}</td>
    </tr>`).join('')}
    </tbody>
  </table></div>`;
}

function toggleEmpSetup(){
  playClick();
  const s=document.getElementById('emp-setup');
  s.style.display=s.style.display==='none'?'block':'none';
}

function addEmployee(){
  const name=document.getElementById('new-emp-name').value.trim();
  const roleKey=document.getElementById('new-emp-role').value;
  const salary=parseFloat(document.getElementById('new-emp-salary').value)||15000;
  if(!name){playError();notify('Enter employee name!','error');return;}
  const emps=getEmployees();
  const roleMap={attendant:{role:'Fuel Attendant',isAttendant:true,isCashier:false,isAdmin:false},cashier:{role:'Cashier',isAttendant:false,isCashier:true,isAdmin:false}};
  const rm=roleMap[roleKey];
  emps.push({id:'emp-'+Date.now(),name,role:rm.role,roleKey,...rm,salary,present:true});
  saveEmployees(emps);
  document.getElementById('new-emp-name').value='';
  document.getElementById('new-emp-salary').value='';
  buildUserSwitcher();
  playSuccess();notify(`✅ Employee ${name} added!`,'success');renderEmployees();
}

function deleteEmployee(id){
  if(!confirm('Delete this employee?'))return;
  const emps=getEmployees().filter(e=>e.id!==id);
  saveEmployees(emps);
  buildUserSwitcher();playSuccess();notify('Employee deleted!','warn');renderEmployees();
}

/* ═══════════════════════════════════
   NOTICES
═══════════════════════════════════ */
function renderNotices(){
  if(isAdminOrCashier(currentUser)){
    document.getElementById('notice-post').style.display='block';
  } else {
    document.getElementById('notice-post').style.display='none';
  }
  const notices=getNotices().slice(-50).reverse();
  const feed=document.getElementById('notices-feed');
  if(!notices.length){feed.innerHTML='<div class="empty-state"><div class="empty-icon">📢</div><p>No announcements yet</p></div>';return;}
  feed.innerHTML=notices.map(n=>`
    <div class="notice-board">
      <div class="notice-time">${n.date} ${n.time}</div>
      <div class="notice-msg">${n.message}</div>
      <div class="notice-by">By: ${n.postedBy}${isAdminOrCashier(currentUser)?` <button class="btn btn-sm btn-danger" style="float:right" onclick="deleteNotice('${n.id}')">Delete</button>`:''}</div>
    </div>`).join('');
}

function postNotice(){
  const msg=document.getElementById('new-notice').value.trim();
  if(!msg){playError();notify('Enter announcement!','error');return;}
  const notices=getNotices(), now=new Date();
  notices.push({id:'notice-'+Date.now(),date:getToday(),time:now.toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}),message:msg,postedBy:currentUser.name});
  saveNotices(notices);
  document.getElementById('new-notice').value='';
  playSuccess();notify('✅ Announcement posted!','success');renderNotices();
}

function deleteNotice(id){
  if(!confirm('Delete?'))return;
  saveNotices(getNotices().filter(n=>n.id!==id));
  playSuccess();notify('Notice deleted!','warn');renderNotices();
}

/* ═══════════════════════════════════
   REPORTS
═══════════════════════════════════ */
function setReportPeriod(period,el){
  playClick();reportPeriod=period;
  document.querySelectorAll('.report-tab').forEach(t=>t.classList.remove('active'));
  if(el)el.classList.add('active');
  renderReports();
}

function renderReports(){
  const txns=getTxns(), today=getToday();
  let filtered;
  if(reportPeriod==='today') filtered=txns.filter(t=>t.date===today);
  else if(reportPeriod==='week') filtered=txns.filter(t=>isWithinDays(t.date,7));
  else if(reportPeriod==='month') filtered=txns.filter(t=>isWithinDays(t.date,30));
  else filtered=txns;
  const totalRev=filtered.reduce((s,t)=>s+t.amount,0);
  const totalGst=filtered.reduce((s,t)=>s+(t.amount*0.18),0);
  const petrolT=filtered.filter(t=>t.fuel==='petrol');
  const dieselT=filtered.filter(t=>t.fuel==='diesel');
  const totalComm=filtered.reduce((s,t)=>s+(t.fuel==='petrol'?t.liters*4:t.liters*3),0);
  document.getElementById('report-stats').innerHTML=`
    <div class="stat-card green"><div class="stat-label">Total Revenue</div><div class="stat-value">₹${fmt(totalRev)}</div><div class="stat-sub">${filtered.length} txns</div></div>
    <div class="stat-card blue"><div class="stat-label">GST 18%</div><div class="stat-value">₹${fmt(totalGst)}</div><div class="stat-sub">Tax collected</div></div>
    <div class="stat-card purple"><div class="stat-label">Commission</div><div class="stat-value">₹${fmt(totalComm)}</div><div class="stat-sub">Owner earnings</div></div>
    <div class="stat-card petrol"><div class="stat-label">Petrol Revenue</div><div class="stat-value">₹${fmt(petrolT.reduce((s,t)=>s+t.amount,0))}</div><div class="stat-sub">${petrolT.length} txns</div></div>
    <div class="stat-card diesel"><div class="stat-label">Diesel Revenue</div><div class="stat-value">₹${fmt(dieselT.reduce((s,t)=>s+t.amount,0))}</div><div class="stat-sub">${dieselT.length} txns</div></div>
  `;
  const tbody=document.getElementById('report-tbody');
  if(!filtered.length){tbody.innerHTML='<tr><td colspan="8"><div class="empty-state">No transactions</div></td></tr>';return;}
  tbody.innerHTML=filtered.map(t=>`<tr>
    <td style="font-size:11px">${t.date}</td>
    <td style="font-family:'DM Mono',monospace;font-size:11px">${t.time}</td>
    <td><span class="badge badge-${t.fuel}">${t.fuel==='petrol'?'⛽ Petrol':'🛢️ Diesel'}</span></td>
    <td style="color:var(--green);font-weight:600">₹${fmt(t.amount)}</td>
    <td style="color:var(--blue)">₹${fmt(t.amount*0.18)}</td>
    <td style="font-family:'DM Mono',monospace">${t.liters.toFixed(2)} L</td>
    <td style="color:var(--text-muted)">₹${t.price.toFixed(2)}</td>
    <td>${t.employee}</td>
  </tr>`).join('');
}

function exportToExcel(){
  if(!window.XLSX){notify('Excel library not loaded!','error');return;}
  const txns=getTxns(), today=getToday();
  let filtered;
  if(reportPeriod==='today') filtered=txns.filter(t=>t.date===today);
  else if(reportPeriod==='week') filtered=txns.filter(t=>isWithinDays(t.date,7));
  else if(reportPeriod==='month') filtered=txns.filter(t=>isWithinDays(t.date,30));
  else filtered=txns;
  if(!filtered.length){notify('No data to export!','warn');return;}
  const data=filtered.map(t=>({'Date':t.date,'Time':t.time,'Fuel Type':t.fuel==='petrol'?'Petrol':'Diesel','Amount (₹)':t.amount,'GST 18% (₹)':t.amount*0.18,'Total with GST (₹)':t.amount*1.18,'Liters':parseFloat(t.liters.toFixed(2)),'Price/L (₹)':parseFloat(t.price.toFixed(2)),'Commission (₹)':t.fuel==='petrol'?t.liters*4:t.liters*3,'Employee':t.employee}));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Transactions');
  XLSX.writeFile(wb,`vignesh_report_${reportPeriod}_${today}.xlsx`);
  playSuccess();notify(`✅ Exported ${filtered.length} transactions!`,'success');
}

/* ═══════════════════════════════════
   SETTINGS
═══════════════════════════════════ */
function loadSettingsInputs(){
  const prices=getPrices();
  document.getElementById('set-petrol-price').value=prices.petrol;
  document.getElementById('set-diesel-price').value=prices.diesel;
  setT('api-status','🔧 Configure API endpoint in Settings');
}

function updatePrices(){
  const p=parseFloat(document.getElementById('set-petrol-price').value)||DEFAULTS.petrolPrice;
  const d=parseFloat(document.getElementById('set-diesel-price').value)||DEFAULTS.dieselPrice;
  savePrices({petrol:p,diesel:d});
  const history=getPriceHistory();
  history.push({date:getToday(),time:new Date().toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}),petrol:p,diesel:d});
  savePriceHistory(history);
  playSuccess();notify(`⛽ Petrol: ₹${p.toFixed(2)}/L | 🛢️ Diesel: ₹${d.toFixed(2)}/L`,'success');
  buildRateHistoryChart();refreshDashboard();
}

function addStock(){
  const p=parseFloat(document.getElementById('add-petrol').value)||0;
  const d=parseFloat(document.getElementById('add-diesel').value)||0;
  if(p===0&&d===0){notify('Enter liters!','error');return;}
  const stock=getStock();
  stock.petrol=Math.min(8000,stock.petrol+p);
  stock.diesel=Math.min(12000,stock.diesel+d);
  saveStock(stock);
  document.getElementById('add-petrol').value='';
  document.getElementById('add-diesel').value='';
  playSuccess();notify(`✅ Added ⛽${p}L Petrol + 🛢️${d}L Diesel!`,'success');refreshDashboard();
}

function resetStock(){
  saveStock({petrol:8000,diesel:12000});
  playSuccess();notify('Stock reset to maximum!','success');refreshDashboard();
}

function clearAllData(){
  saveTxns([]);saveStock({petrol:8000,diesel:12000});
  playSuccess();notify('All data cleared!','warn');refreshDashboard();
}

function exportData(){
  const data={transactions:getTxns(),stock:getStock(),prices:getPrices(),bank:getBank(),exportDate:new Date().toISOString()};
  const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download=`vignesh-backup-${getToday()}.json`;
  document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  notify('Data exported!','success');
}

function fetchLivePrices(){
  setT('api-status','⏳ Fetching...');
  fetch('https://api.example.com/fuel-prices').then(r=>r.json()).then(()=>{setT('api-status','✅ API Connected');notify('Live prices fetched','info');}).catch(()=>{setT('api-status','⚠️ Unable to fetch (Configure API)');notify('Add your own fuel price API endpoint','warn');});
}

/* ═══════════════════════════════════
   CHARTS
═══════════════════════════════════ */
function destroyChart(id){if(charts[id]){charts[id].destroy();delete charts[id];}}

function buildRevenueChart(){
  destroyChart('chart-revenue');
  const txns=getTxns();
  const days=[];
  for(let i=6;i>=0;i--){
    const d=new Date(Date.now()-i*86400000);
    const ds=d.toISOString().split('T')[0];
    const rev=txns.filter(t=>t.date===ds).reduce((s,t)=>s+t.amount,0);
    days.push({lbl:d.toLocaleDateString('en-IN',{weekday:'short',day:'numeric'}),rev});
  }
  const ctx=document.getElementById('chart-revenue');
  if(!ctx)return;
  charts['chart-revenue']=new Chart(ctx,{type:'line',data:{labels:days.map(d=>d.lbl),datasets:[{label:'Revenue (₹)',data:days.map(d=>d.rev),borderColor:'#22d3ee',backgroundColor:'rgba(34,211,238,.1)',fill:true,tension:.4,pointBackgroundColor:'#22d3ee',pointRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{callback:v=>'₹'+v}},x:{grid:{display:false}}}}});
}

function buildRateHistoryChart(){
  destroyChart('chart-rate-history');
  const history=getPriceHistory().slice(-30);
  const ctx=document.getElementById('chart-rate-history');
  if(!ctx)return;
  if(!history.length){
    ctx.parentElement.innerHTML='<div style="color:var(--text-muted);text-align:center;padding:20px">No price history yet</div>';return;
  }
  charts['chart-rate-history']=new Chart(ctx,{type:'line',data:{labels:history.map((h,i)=>i%3===0?h.date:''),datasets:[{label:'⛽ Petrol (₹/L)',data:history.map(h=>h.petrol),borderColor:'#22d3ee',backgroundColor:'rgba(34,211,238,.05)',fill:true,tension:.3,pointRadius:3},{label:'🛢️ Diesel (₹/L)',data:history.map(h=>h.diesel),borderColor:'#f97316',backgroundColor:'rgba(249,115,22,.05)',fill:true,tension:.3,pointRadius:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,position:'bottom'}},scales:{y:{beginAtZero:false,ticks:{callback:v=>'₹'+v.toFixed(0)}},x:{grid:{display:false}}}}});
}

/* ═══════════════════════════════════
   DEMO DATA
═══════════════════════════════════ */
function generateDemoData(){
  const prices=getPrices(), txns=getTxns();
  const names=['Rajesh Kumar','Suresh Babu','Priya Reddy'];
  const fuels=['petrol','diesel','petrol','petrol','diesel','petrol'];
  for(let d=0;d<7;d++){
    const date=new Date(Date.now()-d*86400000).toISOString().split('T')[0];
    const count=Math.floor(Math.random()*12)+8;
    for(let i=0;i<count;i++){
      const fuel=fuels[i%fuels.length];
      const price=fuel==='petrol'?prices.petrol:prices.diesel;
      const amount=[100,200,300,500,1000,1500,2000][Math.floor(Math.random()*7)];
      const liters=amount/price;
      const h=Math.floor(Math.random()*16)+6, m=Math.floor(Math.random()*60);
      txns.push({date,time:`${String(h%12||12).padStart(2,'0')}:${String(m).padStart(2,'0')} ${h<12?'AM':'PM'}`,fuel,amount,liters,price,employee:names[Math.floor(Math.random()*names.length)],branch:currentBranch});
    }
  }
  saveTxns(txns);
  // Also add commission to bank
  const bank=getBank();
  const comm=txns.reduce((s,t)=>s+(t.fuel==='petrol'?t.liters*4:t.liters*3),0)-( getBank().totalEarnings||0);
  bank.totalEarnings=(bank.totalEarnings||0)+Math.max(0,comm);
  bank.balance=(bank.balance||0)+Math.max(0,comm);
  saveBank(bank);
  playSuccess();notify('⚡ Demo data added (7 days)!','success');refreshDashboard();
}

/* ═══════════════════════════════════
   INIT
═══════════════════════════════════ */
function init(){
  // Load branch
  const branches=getBranches();
  currentBranch=branches[0]?.id||'br-01';
  
  startClock();
  buildOutletSelector();

  // Set current user from this branch's employees
  const emps=getEmployees();
  const admin=emps.find(e=>e.isAdmin)||emps[0];
  if(admin) currentUser={id:admin.id,name:admin.name,isAdmin:admin.isAdmin,isCashier:!!admin.isCashier,isAttendant:!!admin.isAttendant};

  buildUserSwitcher();
  buildNavTabs();
  showPage('dashboard');
}

init();
</script>
</body>
</html>
