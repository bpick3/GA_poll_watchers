export function daysUntil(dateStr) {
  const target = new Date(dateStr + 'T00:00:00');
  const now = new Date();
  const diffMs = target.setHours(0,0,0,0) - new Date().setHours(0,0,0,0);
  return Math.round(diffMs / 86400000);
}

export function countdownLabel(dateStr) {
  const d = daysUntil(dateStr);
  if (d > 0) return `${d} day${d === 1 ? '' : 's'}`;
  if (d === 0) return 'Today!';
  return `${-d} day${d === -1 ? '' : 's'} ago`;
}

export function nameOf(people, id) {
  if (!id) return '';
  const p = (people || []).find(p => p.id === id);
  return p ? p.name : '?';
}

export function initials(name) {
  if (!name) return '?';
  return name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
}

export function csvToNames(csv, people) {
  if (!csv) return [];
  return csv.split(',').filter(Boolean).map(id => nameOf(people, id));
}
