import { addDays, parseISODate, toISODateLocal } from './DateUtils.js';

function normalizeDateISO(raw){
  const dt = parseISODate(raw);
  return dt ? toISODateLocal(dt) : null;
}

export function computeDailyLoad(data, employeeID, dateISO){
  const normalized = dateISO;
  const events = (data.courses || []).filter((c) => {
    if (String(c.EmployeeID) !== String(employeeID)) return false;
    return (c.Dates || []).some((d) => normalizeDateISO(d) === normalized);
  });

  const authorities = [...new Set(events.map((e) => String(e.Authority || '').trim()).filter(Boolean))];
  return { eventsCount: events.length, authorities, events };
}

export function computeWeeklyLoad(data, employeeID, startDateISO){
  const start = parseISODate(startDateISO);
  const days = new Set();
  let coursesCount = 0;

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    let countedCourse = false;
    for (const raw of (c.Dates || [])) {
      const d = parseISODate(raw);
      if (!d) continue;
      const diff = Math.floor((d - start) / 86400000);
      if (diff >= 0 && diff <= 6) {
        days.add(toISODateLocal(d));
        countedCourse = true;
      }
    }
    if (countedCourse) coursesCount++;
  }

  return { workDays: days.size, coursesCount };
}

export function computeMonthlyLoad(data, employeeID, dateISO){
  const base = parseISODate(dateISO);
  const month = base.getMonth();
  const year = base.getFullYear();
  const days = new Set();
  let coursesCount = 0;

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    let countedCourse = false;
    for (const raw of (c.Dates || [])) {
      const d = parseISODate(raw);
      if (!d) continue;
      if (d.getMonth() === month && d.getFullYear() === year) {
        days.add(toISODateLocal(d));
        countedCourse = true;
      }
    }
    if (countedCourse) coursesCount++;
  }

  return { workDays: days.size, coursesCount };
}
