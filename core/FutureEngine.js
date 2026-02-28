import { addDays, parseISODate, toISODateLocal } from './DateUtils.js';

export function computeFutureLoad(data, employeeID, dateISO, targetAuthority){
  const start = parseISODate(dateISO);
  const end = addDays(start, 60);
  const targetDow = start.getDay();
  let futureCourses = 0;
  const days = new Set();
  let hasGeoConflictSameWeekday = false;

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    let courseInRange = false;
    for (const raw of (c.Dates || [])) {
      const d = parseISODate(raw);
      if (!d || d <= start || d > end) continue;
      days.add(toISODateLocal(d));
      courseInRange = true;

      if (d.getDay() === targetDow) {
        const authority = String(c.Authority || '').trim();
        if (authority && authority !== String(targetAuthority).trim()) {
          hasGeoConflictSameWeekday = true;
        }
      }
    }
    if (courseInRange) futureCourses++;
  }

  return {
    futureCourses,
    futureWorkDays: days.size,
    hasGeoConflictSameWeekday
  };
}
