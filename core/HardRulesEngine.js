export function validateHardRules(context){
  const { freeSlot, daily, distHome, interAuthorityKm, travel } = context;

  if (!freeSlot) return { ok: false, reason: 'אין חלון זמן' };
  if ((daily?.eventsCount || 0) > 3) return { ok: false, reason: 'יותר מ־3 אירועים ביום' };
  if (distHome === null || distHome > 40) return { ok: false, reason: 'distHome > 40' };
  if ((interAuthorityKm || 0) > 45) return { ok: false, reason: 'קפיצה בין רשויות באותו יום > 45' };
  if (travel && travel.legs > 0 && travel.travelMin <= 0) return { ok: false, reason: 'אין זמן מעבר אמיתי' };
  if (context.future?.hasGeoConflictSameWeekday) return { ok: false, reason: 'קונפליקט עתידי 60 יום באותו יום בשבוע' };

  return { ok: true };
}
