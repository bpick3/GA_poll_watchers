import { useIdentity } from '../identity';

export default function RosterPicker({ peopleData, settings }) {
  const identity = useIdentity();
  const people = peopleData.data || [];
  const tripName = (settings?.data || {}).tripName || 'the trip';

  return (
    <div className="roster-picker">
      <div className="hero-emoji">🍂🏔️🔥</div>
      <h1>{tripName}</h1>
      <p>Who's this? Tap your name to get in.</p>
      <div className="roster-grid">
        {people.map(p => (
          <button key={p.id} className={p.isGuestOfHonor ? 'goh' : ''} onClick={() => identity.setPerson(p.id, p.name)}>
            {p.name}{p.isGuestOfHonor ? ' 🎂' : ''}{p.status === 'maybe' ? ' (maybe)' : ''}
          </button>
        ))}
        {people.length === 0 && <p>Loading roster…</p>}
      </div>
    </div>
  );
}
