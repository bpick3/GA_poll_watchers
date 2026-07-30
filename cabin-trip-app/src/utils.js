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

// Opens the device's native share sheet so the person picks how to send it
// (Messages, WhatsApp, email, etc.) — falls back to the phone's default
// texting app, then to clipboard on desktop where neither is available.
export async function shareText(text, title) {
  if (navigator.share) {
    try {
      await navigator.share({ title, text });
      return 'shared';
    } catch (e) {
      if (e.name === 'AbortError') return 'cancelled';
      // fall through to the sms: fallback below
    }
  }
  const isMobile = /iphone|ipad|ipod|android/i.test(navigator.userAgent);
  if (isMobile) {
    window.location.href = `sms:?&body=${encodeURIComponent(text)}`;
    return 'sms-fallback';
  }
  if (navigator.clipboard) {
    await navigator.clipboard.writeText(text);
    return 'copied';
  }
  return 'unsupported';
}
