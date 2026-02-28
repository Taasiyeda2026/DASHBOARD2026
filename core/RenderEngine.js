export function renderSuggestions(list){
  const container = document.getElementById('resultsList');
  container.innerHTML = '';

  if (!list.length) {
    container.innerHTML = '<div class="loading-placeholder muted">לא נמצאו שיבוצים מתאימים בטווח (מהיום +7 ועד היום +60, א׳–ה׳).</div>';
    return;
  }

  const medals = ['🥇', '🥈', '🥉'];
  list.forEach((c, index) => {
    const card = document.createElement('div');
    card.className = 'result-card';
    card.innerHTML = `
      <div class="item-header">
        <div>
          <div class="result-date">${medals[index] || '🏅'} ${c.name}</div>
          <div class="item-sub">${c.dateISO} | ${c.start}–${c.end}</div>
        </div>
      </div>
      <div style="color:#555;">
        ${c.dateISO || c.date} | ${c.start}–${c.end}
      </div>
      <div style="font-weight:600;">
        ${Number(c.distHome).toFixed(1)} ק"מ
      </div>
    `;
    container.appendChild(card);
  });
}

