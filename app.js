const SALT = "MY_SECRET_2026";
let rawData = [];
let userRole = null;
window.currentUserRole = null;

function fitToScreenIfInstructor(){
  if(window.currentUserRole !== 'instructor') return;

  const wrapper = document.getElementById('app-wrapper');
  if(!wrapper) return;

  if(window.innerWidth <= 800){
    wrapper.style.transform = '';
    return;
  }

  const screenHeight = window.innerHeight;
  const contentHeight = wrapper.scrollHeight;
  const scale = Math.min(1, screenHeight / contentHeight);

  wrapper.style.transform = `scale(${scale})`;
}

window.addEventListener('load', fitToScreenIfInstructor);
window.addEventListener('resize', fitToScreenIfInstructor);

async function sha256(text){
  const encoder = new TextEncoder();
  const data = encoder.encode(text);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b=>b.toString(16).padStart(2,'0')).join('');
}

function showLogin(){
  document.getElementById('loginScreen').innerHTML = `
    <div style="display:flex;flex-direction:column;min-height:100dvh;min-height:100vh;">
      <header style="flex-shrink:0;">
        <div class="header-inner">
          <img src="logo.png" class="app-logo" onerror="this.style.display='none'" />
          <strong>דשבורד פעילויות</strong>
        </div>
      </header>
      <div class="login-page" style="flex:1;min-height:0;">
        <div class="login-bg-shapes">
          <div class="login-shape login-shape-1"></div>
          <div class="login-shape login-shape-2"></div>
          <div class="login-shape login-shape-3"></div>
        </div>
        <div class="login-card">
          <div class="login-logo-wrap">
            <img src="logo.png" class="login-logo" onerror="this.style.display='none'" />
          </div>
          <div class="login-title">כניסה למערכת</div>
          <div class="login-subtitle">דשבורד פעילויות</div>
          <div class="login-form">
            <div class="login-field">
              <label class="login-label">מספר עובד</label>
              <input id="empId" class="login-input" placeholder="הכנס מספר עובד" autocomplete="username" />
            </div>
            <div class="login-field">
              <label class="login-label">קוד אישי</label>
              <input id="empCode" class="login-input" type="password" placeholder="הכנס קוד אישי" autocomplete="current-password" />
            </div>
            <div class="login-remember">
              <input type="checkbox" id="rememberMe" class="remember-checkbox" />
              <label for="rememberMe" class="remember-label">זכור אותי</label>
            </div>
            <button class="login-btn" onclick="login()">
              <span class="login-btn-text">כניסה</span>
              <span class="login-btn-icon">←</span>
            </button>
            <div id="loginError" class="login-error"></div>
          </div>
        </div>
      </div>
    </div>
  `;
  document.getElementById('empId').focus();
  document.getElementById('empCode').addEventListener('keydown', e => { if(e.key === 'Enter') login(); });
  document.getElementById('empId').addEventListener('keydown', e => {
    if(e.key === 'Enter') {
      const code = document.getElementById('empCode').value.trim();
      if(code) login();
      else document.getElementById('empCode').focus();
    }
  });
}

function setLoginError(message){
  const el = document.getElementById('loginError');
  if(el) el.textContent = message;
}

function showLoader(){
  const existing = document.getElementById('loaderOverlay');
  if(existing) return;
  const loader = document.createElement('div');
  loader.id = 'loaderOverlay';
  loader.style.cssText = 'position:fixed;inset:0;background:rgba(255,255,255,0.7);display:flex;align-items:center;justify-content:center;z-index:9999;font-weight:700';
  loader.textContent = 'טוען נתונים...';
  document.body.appendChild(loader);
}

function hideLoader(){
  document.getElementById('loaderOverlay')?.remove();
}

async function fetchJsonWithErrors(path){
  const res = await fetch(path);
  if(res.status === 404){
    throw new Error('not_found');
  }
  if(!res.ok){
    throw new Error('http_error');
  }

  try{
    return await res.json();
  }catch(_e){
    throw new Error('bad_json');
  }
}

async function login(){
  setLoginError('');
  const id = document.getElementById('empId').value.trim();
  const code = document.getElementById('empCode').value.trim();
  if(!id || !code) { setLoginError('יש להזין מספר עובד וקוד אישי.'); return; }

  const hash = await sha256(id + code + SALT);
  const remember = document.getElementById('rememberMe')?.checked ?? false;
  sessionStorage.setItem('dash_empId', id);
  showLoader();

  try{
    try{
      const json = await fetchJsonWithErrors(`./data/instructors/${hash}.json`);
      startApp(json, 'instructor', hash, remember);
      return;
    }catch(e){
      if(e.message !== 'not_found') throw e;
    }

    try{
      const json = await fetchJsonWithErrors(`./data/admins/${hash}.json`);
      startApp(json, 'admin', hash, remember);
      return;
    }catch(e){
      if(e.message === 'not_found'){
        setLoginError('פרטי הכניסה שגויים או הקובץ לא נמצא (404).');
        return;
      }
      throw e;
    }
  }catch(e){
    if(e.message === 'bad_json'){
      setLoginError('קובץ הנתונים פגום (JSON לא תקין).');
    }else{
      setLoginError('שגיאה בטעינת הנתונים. נסו שוב.');
    }
  }finally{
    hideLoader();
  }
}

function startApp(jsonData, role, hash, remember){
  const records = Array.isArray(jsonData) ? jsonData : (jsonData.data || []);
  const name = !Array.isArray(jsonData) ? (jsonData.name || '') : '';

  // Detect dual role (manager + instructor)
  const effectiveRole = (!Array.isArray(jsonData) && jsonData.role === 'both') ? 'both' : role;
  if(effectiveRole === 'both'){
    window.allAdminData = records;
    window.personalData = (!Array.isArray(jsonData) && jsonData.personalData) ? jsonData.personalData : [];
  }

  rawData = records;
  userRole = effectiveRole;
  window.currentUserRole = effectiveRole;
  document.body.dataset.role = effectiveRole;

  if(remember !== null){
    const store = remember ? localStorage : sessionStorage;
    store.setItem('dash_hash', hash);
    store.setItem('dash_role', effectiveRole);
    if(name) store.setItem('dash_name', name);
  }
  window.EmployeeID = sessionStorage.getItem('dash_empId') || '';

  document.getElementById('loginScreen').style.display='none';
  document.getElementById('app').style.display='flex';

  // החזרת גלילה לראש הדף לאחר כניסה (במובייל הדפדפן יכול להישאר גלול למטה)
  window.scrollTo(0, 0);
  document.documentElement.scrollTop = 0;
  document.body.scrollTop = 0;

  const logoutBtn = document.getElementById('logoutBtn');
  if(logoutBtn){
    logoutBtn.onclick = ()=>{
      sessionStorage.clear();
      localStorage.removeItem('dash_hash');
      localStorage.removeItem('dash_role');
      localStorage.removeItem('dash_name');
      location.reload();
    };
  }

  fitToScreenIfInstructor();
  loadDashboard();
}

function loadDashboard(){
  const script = document.createElement('script');
  script.src = 'dashboard.js';
  document.body.appendChild(script);
}

async function resumeSession(){
  const hash = localStorage.getItem('dash_hash') || sessionStorage.getItem('dash_hash');
  const role = localStorage.getItem('dash_role') || sessionStorage.getItem('dash_role');
  if(!hash || !role) return false;

  showLoader();
  try{
    let json;
    if(role === 'admin'){
      json = await fetchJsonWithErrors(`./data/admins/${hash}.json`);
    } else if(role === 'both'){
      // 'both' files can be in either instructors (existing) or admins (generator)
      try{
        json = await fetchJsonWithErrors(`./data/instructors/${hash}.json`);
      }catch(e){
        if(e.message !== 'not_found') throw e;
        json = await fetchJsonWithErrors(`./data/admins/${hash}.json`);
      }
    } else {
      json = await fetchJsonWithErrors(`./data/instructors/${hash}.json`);
    }
    startApp(json, role, hash, null);
    return true;
  }catch(_e){
    sessionStorage.clear();
    localStorage.removeItem('dash_hash');
    localStorage.removeItem('dash_role');
    localStorage.removeItem('dash_name');
    return false;
  }finally{
    hideLoader();
  }
}

(async ()=>{
  showLogin();
  await resumeSession();
})();
