export function scoreCandidate(context){
  const { distHome, daily, weekly, monthly, future, alreadyInAuthority } = context;
  let score = 1000;

  score -= (distHome || 0) * 6;
  score -= (daily.travelKm || 0) * 3;
  score -= (daily.eventsCount || 0) * 40;
  score -= (weekly.workDays || 0) * 25;
  score -= (monthly.coursesCount || 0) * 5;
  score -= (future.futureCourses || 0) * 8;
  score -= (future.futureWorkDays || 0) * 15;

  if (future.hasGeoConflictSameWeekday) score -= 250;
  if ((daily.eventsCount || 0) === 0) score += 120;
  if (alreadyInAuthority) score += 180;
  if ((distHome || 0) <= 15) score += 80;

  let quality = 'גבולי';
  if (score >= 850) quality = 'אידאלי';
  else if (score >= 700) quality = 'סביר';

  return { score: Math.round(score), quality };
}
