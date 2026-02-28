export function buildExplanation(candidate){
  return [
    `מרחק מהבית: ${candidate.distHome.toFixed(1)} ק"מ`,
    `זמן נסיעה יומי: ${Math.round(candidate.daily.travelMin)} דק׳`,
    `אירועים היום: ${candidate.daily.eventsCount}`,
    `עומס שבועי: ${candidate.weekly.workDays} ימי עבודה`,
    `עומס חודשי: ${candidate.monthly.coursesCount} קורסים`,
    `עומס 60 יום: ${candidate.future.futureCourses} קורסים / ${candidate.future.futureWorkDays} ימים`,
    `קונפליקט עתידי: ${candidate.future.hasGeoConflictSameWeekday ? 'כן' : 'לא'}`,
    `ציון: ${candidate.score} (${candidate.quality})`
  ];
}
