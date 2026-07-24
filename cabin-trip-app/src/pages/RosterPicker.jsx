import { useIdentity } from '../identity';

export default function RosterPicker({ peopleData }) {
  const identity = useIdentity();
  const people = peopleData.data || [];

  return (
    <div className="roster-picker">
      <div className="hero-emoji">🍂🏔️🔥</div>
      <h1>Cabin Trip 2026</h1>
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
