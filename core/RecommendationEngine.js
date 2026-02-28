const DEBUG = true;

function toISO(date){
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function startOfDay(date){
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function addDays(date, days){
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function parseISODate(raw){
  if (typeof raw !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(raw)) return null;
  const [y, m, d] = raw.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function excelSerialToISO(serial){
  if (typeof serial !== 'number' || !Number.isFinite(serial)) return null;
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const dt = new Date(excelEpoch.getTime() + Math.floor(serial) * 86400000);
  const y = dt.getUTCFullYear();
  const m = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const d = String(dt.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function timeToMin(raw, fallback){
  if (typeof raw === 'string') {
    const m = raw.trim().match(/^(\d{1,2}):(\d{2})$/);
    if (m) return Number(m[1]) * 60 + Number(m[2]);
  }

  if (typeof raw === 'number' && Number.isFinite(raw)) {
    if (raw >= 0 && raw <= 1) return Math.round(raw * 24 * 60);
    return Math.round((raw % 1) * 24 * 60);
  }

  return fallback;
}

function minToTime(min){
  const hh = String(Math.floor(min / 60)).padStart(2, '0');
  const mm = String(min % 60).padStart(2, '0');
  return `${hh}:${mm}`;
}

function normalizeCourseDatesToISO(data){
  return {
    ...data,
    courses: (data.courses || []).map((c) => ({
      ...c,
      Dates: (c.Dates || []).map((d) => {
        if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
        return excelSerialToISO(d);
      }).filter(Boolean)
    }))
  };
}

function getSuggestionDates(){
  const today = startOfDay(new Date());
  const results = [];

  for (let i = 7; i <= 30; i++) {
    const d = addDays(today, i);
    const day = d.getDay();
    if (day >= 0 && day <= 4) results.push(toISO(d));
  }

  return results;
}

function deg2rad(v){
  return (v * Math.PI) / 180;
}

function haversine(lat1, lon1, lat2, lon2){
  const R = 6371;
  const dLat = deg2rad(lat2 - lat1);
  const dLon = deg2rad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(deg2rad(lat1)) * Math.cos(deg2rad(lat2)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);

  return R * (2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function distanceKm(data, fromAuthority, toAuthority){
  if (!fromAuthority || !toAuthority) return null;
  if (fromAuthority === toAuthority) return 0;
  const a = data.authorityLocations?.[fromAuthority];
  const b = data.authorityLocations?.[toAuthority];
  if (!a || !b) return null;
  return haversine(a.lat, a.lng, b.lat, b.lng);
}

function collectEventsByDate(data, employeeID, dateISO){
  const events = [];
  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    if (!Array.isArray(c.Dates) || !c.Dates.includes(dateISO)) continue;

    events.push({
      authority: String(c.Authority || '').trim(),
      school: String(c.School || '').trim(),
      startMin: timeToMin(c.StartTime, 8 * 60),
      endMin: timeToMin(c.EndTime, 15 * 60)
    });
  }

  return events.sort((a, b) => a.startMin - b.startMin);
}

function buildSlot(events, durationMin){
  const workStart = 8 * 60;
  const workEnd = 15 * 60;
  let cursor = workStart;

  for (const ev of events) {
    if (ev.startMin - cursor >= durationMin) {
      return { startMin: cursor, endMin: cursor + durationMin, freeSlot: true };
    }
    cursor = Math.max(cursor, ev.endMin);
  }

  if (workEnd - cursor >= durationMin) {
    return { startMin: cursor, endMin: cursor + durationMin, freeSlot: true };
  }

  return { startMin: null, endMin: null, freeSlot: false };
}

function validateTravelChain(data, events){
  const sorted = [...events].sort((a, b) => a.startMin - b.startMin);

  for (let i = 0; i < sorted.length - 1; i++) {
    const current = sorted[i];
    const next = sorted[i + 1];

    if (current.endMin > next.startMin) return { ok: false, reason: 'overlap' };
    if (current.school && next.school && current.school === next.school) continue;

    const kmRaw = distanceKm(data, current.authority, next.authority);
    const km = kmRaw === null ? Infinity : kmRaw;

    if (current.authority === next.authority) {
      if (km > 15) return { ok: false, reason: 'distance_same_authority' };
    } else if (km > 30) {
      return { ok: false, reason: 'distance_cross_authority' };
    }

    const gap = next.startMin - current.endMin;
    const travelRequired = Math.ceil(km);
    if (gap < travelRequired) return { ok: false, reason: 'travel_time' };
  }

  return { ok: true, reason: '' };
}

function computeWeeklyWorkDays(data, employeeID, dateISO){
  const base = parseISODate(dateISO);
  if (!base) return 0;

  const day = base.getDay();
  const weekStart = addDays(base, -day);
  const weekEnd = addDays(weekStart, 6);
  const worked = new Set();

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    for (const d of (c.Dates || [])) {
      const dt = parseISODate(d);
      if (!dt) continue;
      if (dt >= weekStart && dt <= weekEnd) worked.add(d);
    }
  }

  worked.add(dateISO);
  return worked.size;
}

function computeMonthlyCourses(data, employeeID, dateISO){
  const base = parseISODate(dateISO);
  if (!base) return 0;

  const y = base.getFullYear();
  const m = base.getMonth();
  let count = 0;

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;
    for (const d of (c.Dates || [])) {
      const dt = parseISODate(d);
      if (!dt) continue;
      if (dt.getFullYear() === y && dt.getMonth() === m) count += 1;
    }
  }

  return count + 1;
}

function computeFutureStats(data, employeeID, candidateDateISO, targetAuthority){
  const today = startOfDay(new Date());
  const horizon = addDays(today, 60);
  const candidateDate = parseISODate(candidateDateISO);
  const candidateWeekday = candidateDate ? candidateDate.getDay() : null;

  let futureCourses = 0;
  const workDaysSet = new Set();
  let weekdayDistanceConflict = false;

  for (const c of (data.courses || [])) {
    if (String(c.EmployeeID) !== String(employeeID)) continue;

    for (const d of (c.Dates || [])) {
      const dt = parseISODate(d);
      if (!dt || dt < today || dt > horizon) continue;

      futureCourses += 1;
      workDaysSet.add(d);

      if (candidateWeekday !== null && dt.getDay() === candidateWeekday) {
        const km = distanceKm(data, String(c.Authority || '').trim(), targetAuthority);
        if (km !== null && km > 30) {
          weekdayDistanceConflict = true;
        }
      }
    }
  }

  return {
    futureCourses,
    futureWorkDays: workDaysSet.size,
    weekdayDistanceConflict
  };
}

function buildQuality(score){
  if (score >= 1150) return 'A';
  if (score >= 1000) return 'B';
  if (score >= 850) return 'C';
  return 'D';
}

export function buildGlobalRecommendations(data, targetAuthority, durationMin, topN){
  const normalized = normalizeCourseDatesToISO(data);
  const instructors = (normalized.instructors || []).filter((i) => String(i.Role || '').toLowerCase() === 'instructor');
  const suggestionDates = getSuggestionDates();
  const candidates = [];
  const uniqueSlots = new Set();

  const debugStats = {
    rejectedByDateWindow: 0,
    rejectedByWeekday: 0,
    rejectedByTime: 0,
    rejectedByDistance: 0,
    rejectedByTravel: 0,
    rejectedByLoad: 0,
    rejectedByFutureConflict: 0,
    validCandidates: 0
  };

  for (const dateISO of suggestionDates) {
    const dateObj = parseISODate(dateISO);
    if (!dateObj) {
      debugStats.rejectedByDateWindow += instructors.length;
      continue;
    }

    const day = dateObj.getDay();
    if (day < 0 || day > 4) {
      debugStats.rejectedByWeekday += instructors.length;
      continue;
    }

    for (const inst of instructors) {
      const employeeID = String(inst.EmployeeID);
      const eventsToday = collectEventsByDate(normalized, employeeID, dateISO);

      if (eventsToday.length >= 3) {
        debugStats.rejectedByLoad += 1;
        continue;
      }

      const slot = buildSlot(eventsToday, durationMin);
      if (!slot.freeSlot) {
        debugStats.rejectedByTime += 1;
        continue;
      }

      const stagedEvents = [...eventsToday, {
        authority: targetAuthority,
        school: '',
        startMin: slot.startMin,
        endMin: slot.endMin
      }];

      const travelValidation = validateTravelChain(normalized, stagedEvents);
      if (!travelValidation.ok) {
        if (travelValidation.reason.startsWith('distance')) debugStats.rejectedByDistance += 1;
        else debugStats.rejectedByTravel += 1;
        continue;
      }

      const future = computeFutureStats(normalized, employeeID, dateISO, targetAuthority);
      if (future.weekdayDistanceConflict) {
        debugStats.rejectedByFutureConflict += 1;
        continue;
      }

      const homeAuthority = String(inst.HomeAuthority || '').trim();
      const distHomeRaw = distanceKm(normalized, homeAuthority, targetAuthority);
      const distHome = distHomeRaw === null ? 999 : distHomeRaw;

      const weeklyWorkDays = computeWeeklyWorkDays(normalized, employeeID, dateISO);
      const monthlyCourses = computeMonthlyCourses(normalized, employeeID, dateISO);

      let score = 1000;
      score -= distHome * 5;
      score -= eventsToday.length * 40;
      score -= weeklyWorkDays * 25;
      score -= monthlyCourses * 5;
      score -= future.futureCourses * 8;

      if (eventsToday.length === 0) score += 120;
      if (eventsToday.some((e) => e.authority === targetAuthority)) score += 150;
      if (distHome <= 10) score += 80;

      score = Math.max(0, Math.round(score));

      const candidate = {
        employeeID,
        name: inst.Employee || '—',
        dateISO,
        start: minToTime(slot.startMin),
        end: minToTime(slot.endMin),
        distHome,
        score,
        quality: buildQuality(score),
        daily: {
          eventsCount: eventsToday.length,
          travelMin: 0
        },
        weekly: {
          workDays: weeklyWorkDays
        },
        monthly: {
          coursesCount: monthlyCourses
        },
        future
      };

      const uniqueKey = `${employeeID}_${dateISO}_${candidate.start}_${candidate.end}`;
      if (uniqueSlots.has(uniqueKey)) continue;
      uniqueSlots.add(uniqueKey);

      debugStats.validCandidates += 1;
      candidates.push(candidate);
    }
  }

  const bestByInstructor = new Map();
  for (const candidate of candidates) {
    const prev = bestByInstructor.get(candidate.employeeID);
    if (!prev || candidate.score > prev.score) bestByInstructor.set(candidate.employeeID, candidate);
  }

  if (DEBUG) {
    // eslint-disable-next-line no-console
    console.table(debugStats);
  }

  const recommendations = [...bestByInstructor.values()]
    .sort((a, b) => b.score - a.score)
    .slice(0, topN);

  return { recommendations, debugStats };
}
