import { getDistanceKm, estimateTravelMinutes } from './GeoUtils.js';

export function computeDayTravelStats(data, events, candidateAuthority){
  const route = [...events.map((e) => e.authority).filter(Boolean)];
  if (candidateAuthority) route.push(candidateAuthority);

  let travelKm = 0;
  let travelMin = 0;
  for (let i = 1; i < route.length; i++) {
    const km = getDistanceKm(data, route[i - 1], route[i]);
    if (km === null) continue;
    travelKm += km;
    travelMin += estimateTravelMinutes(km);
  }

  return {
    route,
    legs: Math.max(route.length - 1, 0),
    travelKm,
    travelMin
  };
}
