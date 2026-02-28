function pad2(n){ return String(n).padStart(2, '0'); }

export function parseISODate(value){
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  if (typeof value === 'number' && Number.isFinite(value)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dt = new Date(excelEpoch.getTime() + value * 86400000);
    return new Date(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
  }
  if (typeof value !== 'string') return null;
  const m = value.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

export function addDays(date, days){
  const d = parseISODate(date) || new Date(date);
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  x.setDate(x.getDate() + days);
  return x;
}

export function toISODateLocal(date){
  const d = parseISODate(date) || new Date(date);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

export function isWorkdaySunThu(dateObj){
  const d = dateObj.getDay();
  return d >= 0 && d <= 4;
}

export function getSuggestionRangeISO(){
  const today = new Date();
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const start = addDays(base, 7);
  const end = addDays(base, 60);
  const dates = [];
  for (let d = start; d <= end; d = addDays(d, 1)) {
    if (isWorkdaySunThu(d)) dates.push(toISODateLocal(d));
  }
  return dates;
}
